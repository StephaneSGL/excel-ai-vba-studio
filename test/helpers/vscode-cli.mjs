import { readdir, stat } from 'node:fs/promises';
import { dirname, join } from 'node:path';

// Recent VS Code archives place resources under a versioned subdirectory.
// Discover the CLI file without executing/parsing a shell wrapper or guessing a hash.
export async function resolveWindowsCliPath(executable) {
  const directory = dirname(executable);
  const entries = await readdir(directory, { withFileTypes: true });
  const candidates = [join(directory, 'resources/app/out/cli.js'),
    ...entries.filter(entry => entry.isDirectory()).map(entry => join(directory, entry.name, 'resources/app/out/cli.js'))];
  const found = [];
  for (const candidate of candidates) {
    try { if ((await stat(candidate)).isFile()) found.push(candidate); }
    catch (error) { if (error.code !== 'ENOENT' && error.code !== 'ENOTDIR') throw error; }
  }
  if (found.length !== 1) throw new Error(`Expected one VS Code CLI entry point; found ${found.length}.`);
  return found[0];
}
