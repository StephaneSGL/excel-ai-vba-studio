/** Platform support, not a probe for an installed/licensed copy of Excel. */
export interface PlatformCapabilities {
    readonly platform: string;
    readonly portableSpreadsheet: boolean;
    readonly nativeExcel: boolean;
}

export function getPlatformCapabilities(
    platform: string = process.platform,
    architecture: string = process.arch,
): PlatformCapabilities {
    return Object.freeze({
        platform,
        portableSpreadsheet: true,
        nativeExcel: platform === 'win32' && architecture === 'x64',
    });
}

export const PORTABLE_PRESERVATION_MESSAGE =
    'Lecture seule : ce classeur contient des éléments que le mode portable ne peut pas préserver. Ouvrez-le avec Excel Desktop sous Windows pour le modifier.';

export const NATIVE_EXCEL_UNAVAILABLE_MESSAGE =
    'Cette fonction nécessite Excel Desktop sous Windows x64. La grille de tableur portable reste disponible sur cette plateforme.';
