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
