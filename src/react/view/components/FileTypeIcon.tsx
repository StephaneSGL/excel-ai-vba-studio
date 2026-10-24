import { useMemo, useState } from 'react';
import { getFileIconUrl, getFolderIconUrl } from './fileIcon';

interface FileTypeIconProps {
    name?: string;
    isDirectory?: boolean;
    className?: string;
}

import { sanitizeFileIconUrl } from '../../util/sanitizeFileIconUrl';
export { sanitizeFileIconUrl } from '../../util/sanitizeFileIconUrl';

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
