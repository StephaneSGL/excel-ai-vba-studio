import { useMemo, useState } from 'react';
import { IconFile, IconFolder, IconFolderOpen } from '../icons';
import { getFileIconUrl, getFolderIconUrl } from '../fileIcon';

import { sanitizeFileIconUrl } from '../../../util/sanitizeFileIconUrl';
export { sanitizeFileIconUrl } from '../../../util/sanitizeFileIconUrl';

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
