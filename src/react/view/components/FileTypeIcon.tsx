import { useMemo, useState } from 'react';
import { getFileIconUrl, getFolderIconUrl } from './fileIcon';

interface FileTypeIconProps {
    name?: string;
    isDirectory?: boolean;
    className?: string;
}

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

export function FileTypeIcon({ name, isDirectory, className }: FileTypeIconProps) {
    const [failed, setFailed] = useState(false);
    const iconUrl = useMemo(() => {
        if (!name) return null;
        const candidate = isDirectory ? getFolderIconUrl(name) : getFileIconUrl(name);
        return sanitizeFileIconUrl(candidate);
    }, [name, isDirectory]);

    if (!iconUrl || failed) {
        const codicon = isDirectory ? 'folder' : 'file';
        return (
            <span
                className={`codicon codicon-${codicon}${className ? ` ${className}` : ''}`}
                aria-hidden
            />
        );
    }

    return (
        <img
            className={className ?? 'file-type-icon'}
            src={iconUrl}
            alt=""
            draggable={false}
            onError={() => setFailed(true)}
        />
    );
}
