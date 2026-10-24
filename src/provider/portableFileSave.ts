import { createHash, randomUUID } from 'node:crypto';
import { spawn } from 'node:child_process';
import { constants, promises as fs } from 'node:fs';
import type { BigIntStats } from 'node:fs';
import type { FileHandle } from 'node:fs/promises';
import { basename, dirname, extname, join, parse, resolve } from 'node:path';

const MAX_PORTABLE_FILE_BYTES = 128 * 1024 * 1024;
const BACKUP_DIRECTORY = '.excel-ai-vba-backups';
const pendingSaves = new Map<string, Promise<string>>();
const ATTRIBUTE_COPY_TIMEOUT_MS = 10_000;
const MAX_ATTRIBUTE_COPY_STDERR_BYTES = 8 * 1024;

function sameLocalPath(left: string, right: string): boolean {
    // VS Code normalizes drive letters to lowercase; Windows realpath can return
    // uppercase volume names. This is the same directory, not a symlink redirect.
    const normalizeDrive = (value: string) => value.replace(/^[a-z]:/i, drive => drive.toUpperCase());
    return process.platform === 'win32' ? normalizeDrive(left) === normalizeDrive(right) : left === right;
}

function hash(bytes: Uint8Array): string {
    return createHash('sha256').update(bytes).digest('hex');
}

function validateBytes(bytes: Uint8Array): void {
    if (!(bytes instanceof Uint8Array) || bytes.byteLength > MAX_PORTABLE_FILE_BYTES) {
        throw new Error('Spreadsheet save exceeds the 128 MiB safety limit.');
    }
}

async function localParent(filePath: string): Promise<string> {
    const fullPath = resolve(filePath);
    const parent = dirname(fullPath);
    if (containsRecoveryDirectory(fullPath)) {
        throw new Error('Recovery backups cannot be overwritten by the spreadsheet editor.');
    }
    if (process.platform === 'win32') {
        // realpath expands legitimate 8.3 names and corrects directory casing.
        // Inspect the original path first: canonicalization alone would hide a junction.
        await assertWindowsDirectoryChain(parent);
        const canonicalParent = await fs.realpath(parent);
        await assertWindowsDirectoryChain(canonicalParent);
        if (containsRecoveryDirectory(join(canonicalParent, basename(fullPath)))) {
            throw new Error('Recovery backups cannot be overwritten by the spreadsheet editor.');
        }
        return canonicalParent;
    }
    // A different real parent would let a symlink redirect a staged replacement.
    if (!sameLocalPath(await fs.realpath(parent), parent)) {
        throw new Error('Portable saves require a local path without symbolic links.');
    }
    return parent;
}

function containsRecoveryDirectory(filePath: string): boolean {
    return filePath.split(/[\\/]/).some(part => process.platform === 'win32'
        ? part.toLowerCase() === BACKUP_DIRECTORY
        : part === BACKUP_DIRECTORY);
}

async function assertWindowsDirectoryChain(directory: string): Promise<void> {
    const root = parse(directory).root;
    let current = root;
    const segments = directory.slice(root.length).split(/[\\/]/).filter(Boolean);
    // lstat reports Windows junctions as symbolic links. Include all ancestors,
    // even when the immediate parent itself is an ordinary directory.
    for (const segment of ['', ...segments]) {
        if (segment) current = join(current, segment);
        const info = await fs.lstat(current);
        if (!info.isDirectory() || info.isSymbolicLink()) {
            throw new Error('Portable saves require local directories without symbolic links or junctions.');
        }
    }
}

async function readStable(handle: FileHandle): Promise<Buffer> {
    const before = await handle.stat({ bigint: true });
    if (!before.isFile() || before.size > MAX_PORTABLE_FILE_BYTES) {
        throw new Error('Spreadsheet source is not a regular file within the 128 MiB limit.');
    }
    const bytes = Buffer.alloc(Number(before.size));
    let offset = 0;
    while (offset < bytes.length) {
        const result = await handle.read(bytes, offset, bytes.length - offset, offset);
        if (result.bytesRead === 0) throw new Error('Spreadsheet changed while it was being read.');
        offset += result.bytesRead;
    }
    const extra = await handle.read(Buffer.alloc(1), 0, 1, offset);
    const after = await handle.stat({ bigint: true });
    if (extra.bytesRead || !sameSourceState(before, after)) {
        throw new Error('Spreadsheet changed while it was being read.');
    }
    return bytes;
}

function sameSourceState(left: BigIntStats, right: BigIntStats): boolean {
    return left.dev === right.dev && left.ino === right.ino
        && left.size === right.size && left.nlink === right.nlink
        && left.mode === right.mode && left.uid === right.uid && left.gid === right.gid
        && left.mtimeNs === right.mtimeNs && left.ctimeNs === right.ctimeNs;
}

async function preserveLinuxAttributes(source: FileHandle, stagePath: string): Promise<void> {
    // POSIX mode bits expose an ACL's mask, not its owning-group permissions.
    // chmod alone can therefore grant access which the original ACL denied.
    // GNU cp preserves ACLs with mode and copies xattrs without changing bytes.
    await new Promise<void>((resolvePromise, reject) => {
        const child = spawn('/usr/bin/cp', [
            '--attributes-only', '--preserve=mode,ownership,xattr', '--no-target-directory',
            '--', '/proc/self/fd/3', stagePath,
        ], {
            shell: false,
            stdio: ['ignore', 'ignore', 'pipe', source.fd],
        });
        let failed = false;
        let stderrBytes = 0;
        const timeout = setTimeout(() => {
            failed = true;
            child.kill('SIGKILL');
        }, ATTRIBUTE_COPY_TIMEOUT_MS);
        child.on('error', () => { failed = true; });
        child.stderr?.on('data', (chunk: Buffer) => {
            stderrBytes += chunk.byteLength;
            if (stderrBytes > MAX_ATTRIBUTE_COPY_STDERR_BYTES) {
                failed = true;
                child.kill('SIGKILL');
            }
        });
        // Wait for close even after error/timeout: never rename while cp runs.
        child.once('close', (code, signal) => {
            clearTimeout(timeout);
            if (failed || code !== 0 || signal || stderrBytes !== 0) {
                reject(new Error(
                    'Linux save could not preserve filesystem permissions and attributes. '
                    + 'GNU coreutils /usr/bin/cp is required; the original file was not replaced.',
                ));
            } else {
                resolvePromise();
            }
        });
    });
}

async function writeExclusive(filePath: string, bytes: Uint8Array, mode: number): Promise<void> {
    const handle = await fs.open(filePath, 'wx', mode);
    try {
        await handle.writeFile(bytes);
        await handle.sync();
    } finally {
        await handle.close();
    }
}

/** Create without replacing a file selected accidentally in Save As. */
export async function createPortableFile(filePath: string, bytes: Uint8Array): Promise<string> {
    validateBytes(bytes);
    const fullPath = resolve(filePath);
    const parent = await localParent(fullPath);
    const destination = process.platform === 'win32' ? join(parent, basename(fullPath)) : fullPath;
    const staging = await fs.mkdtemp(join(parent, '.excel-ai-save-'));
    const stagePath = join(staging, 'workbook');
    try {
        await writeExclusive(stagePath, bytes, 0o600);
        if (process.platform === 'win32' && !sameLocalPath(await localParent(fullPath), parent)) {
            throw new Error('Save destination changed.');
        }
        if (!sameLocalPath(await fs.realpath(parent), parent)) throw new Error('Save destination changed.');
        // link(2) installs the complete file atomically and fails if the target exists.
        // Unlike rename, this can never clobber a destination created during Save As.
        await fs.link(stagePath, destination);
        return hash(bytes);
    } catch (error) {
        if ((error as NodeJS.ErrnoException).code === 'EEXIST') {
            throw new Error('The destination already exists. Choose a new name to preserve that file.');
        }
        throw error;
    } finally {
        await fs.rm(staging, { recursive: true, force: true });
    }
}

/**
 * Linux staged save tied to the bytes last delivered by the host. Windows uses
 * its existing file writer so NTFS origin streams are not lost by a POSIX rename.
 * The recovery copy preserves the previous bytes even if a later writer races us.
 */
export function savePortableFile(
    filePath: string,
    bytes: Uint8Array,
    expectedSha256: string,
): Promise<string> {
    const key = resolve(filePath);
    const previous = pendingSaves.get(key) ?? Promise.resolve('');
    const current = previous.catch(() => '').then(() =>
        savePortableFileUnlocked(key, bytes, expectedSha256));
    pendingSaves.set(key, current);
    void current.finally(() => {
        if (pendingSaves.get(key) === current) pendingSaves.delete(key);
    }).catch(() => undefined);
    return current;
}

async function savePortableFileUnlocked(
    filePath: string,
    bytes: Uint8Array,
    expectedSha256: string,
): Promise<string> {
    if (process.platform !== 'linux') {
        throw new Error('Staged workbook replacement is qualified only on Linux.');
    }
    validateBytes(bytes);
    if (!/^[a-f0-9]{64}$/.test(expectedSha256)) {
        throw new Error('Reload this spreadsheet before saving: its source snapshot is missing.');
    }
    const fullPath = resolve(filePath);
    const parent = await localParent(fullPath);
    const initial = await fs.lstat(fullPath, { bigint: true });
    if (!initial.isFile() || initial.isSymbolicLink() || initial.nlink !== 1n) {
        throw new Error('Portable saves require a regular file without symbolic or hard links.');
    }
    if ((initial.mode & 0o222n) === 0n) throw new Error('This spreadsheet is read-only.');
    const source = await fs.open(fullPath, constants.O_RDONLY | constants.O_NOFOLLOW);
    let staging: string | undefined;
    let recoveryPath: string | undefined;
    let committed = false;
    try {
        const opened = await source.stat({ bigint: true });
        if (!sameSourceState(opened, initial)) {
            throw new Error('Spreadsheet was replaced before saving. Reload it first.');
        }
        const previousBytes = await readStable(source);
        if (hash(previousBytes) !== expectedSha256) {
            throw new Error('Spreadsheet changed outside VS Code. Reload it before saving.');
        }
        staging = await fs.mkdtemp(join(parent, '.excel-ai-save-'));
        const stagePath = join(staging, 'workbook');
        await writeExclusive(stagePath, bytes, 0o600);
        await preserveLinuxAttributes(source, stagePath);

        const backups = join(parent, BACKUP_DIRECTORY);
        await fs.mkdir(backups, { mode: 0o700 }).catch(error => {
            if ((error as NodeJS.ErrnoException).code !== 'EEXIST') throw error;
        });
        const backupInfo = await fs.lstat(backups);
        if (!backupInfo.isDirectory() || backupInfo.isSymbolicLink()
            || (backupInfo.mode & 0o022) !== 0 || await fs.realpath(backups) !== backups) {
            throw new Error('The recovery directory is not a private local directory.');
        }
        recoveryPath = join(backups, `portable-${Date.now()}-${randomUUID()}${extname(fullPath).slice(0, 16)}`);
        await writeExclusive(recoveryPath, previousBytes, 0o600);

        const currentHash = hash(await readStable(source));
        const currentParent = await fs.realpath(parent);
        const current = await fs.lstat(fullPath, { bigint: true });
        if (!sameSourceState(current, initial) || currentParent !== parent
            || currentHash !== expectedSha256) {
            throw new Error('Spreadsheet changed during saving. Reload it before trying again.');
        }
        await fs.rename(stagePath, fullPath);
        committed = true;
        return hash(bytes);
    } finally {
        await source.close();
        if (staging) await fs.rm(staging, { recursive: true, force: true });
        if (!committed && recoveryPath) await fs.rm(recoveryPath, { force: true });
    }
}
