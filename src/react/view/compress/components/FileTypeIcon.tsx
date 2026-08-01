import { useMemo, useState } from 'react';
import { IconFile, IconFolder, IconFolderOpen } from '../icons';
import { getFileIconUrl, getFolderIconUrl } from '../fileIcon';

const VSCODE_RESOURCE_HOST_SUFFIX = '.vscode-resource.vscode-cdn.net';

export function sanitizeFileIconUrl(iconUrl: string | null): string | null {
    if (!iconUrl) return null;
    try {
        const parsed = new URL(iconUrl);
        const trustedHost = parsed.hostname.endsWith(VSCODE_RESOURCE_HOST_SUFFIX);
        if (
            parsed.protocol !== 'https:'
            || !trustedHost
            || parsed.username
            || parsed.password
            || parsed.port
            || parsed.search
            || parsed.hash
            || !parsed.pathname.toLowerCase().endsWith('.svg')
        ) {
            return null;
        }
        return parsed.href;
    } catch {
        return null;
    }
}

export function FileTypeIcon({ name, isDirectory, expanded }: { name?: string; isDirectory?: boolean; expanded?: boolean }) {
    const [failed, setFailed] = useState(false);
    const iconUrl = useMemo(() => {
        if (!name) return null;
        const candidate = isDirectory ? getFolderIconUrl(name, expanded) : getFileIconUrl(name);
        return sanitizeFileIconUrl(candidate);
    }, [name, isDirectory, expanded]);

    if (!iconUrl || failed) {
        return isDirectory
            ? (expanded ? <IconFolderOpen size={15} /> : <IconFolder size={15} />)
            : <IconFile size={15} />;
    }
    return (
        <img
            className="zip-type-icon"
            src={iconUrl}
            alt=""
            draggable={false}
            onError={() => setFailed(true)}
        />
    );
}
