export interface PopupAnchor {
    x: number;
    y: number;
    /** Toolbar / repo settings: place below cursor when near top of view */
    repoToolbar?: boolean;
    /** Center dialog horizontally on the anchor instead of biasing left */
    centerHorizontally?: boolean;
}

export interface ViewportBounds {
    top: number;
    left: number;
    width: number;
    height: number;
}

export type AnchoredDialogPositionVariant = 'default' | 'merge';

export const COMMIT_DETAIL_POPUP_WIDTH = 400;
export const POPUP_MARGIN = 10;
const MAX_COMMIT_DETAIL_HEIGHT = 600;
const ANCHOR_TOP_OFFSET = 40;
const ANCHOR_LEFT_OFFSET = 32;

export function getViewportBounds(container?: HTMLElement | null): ViewportBounds {
    if (container) {
        const rect = container.getBoundingClientRect();
        return {
            top: rect.top,
            left: rect.left,
            width: Math.max(0, rect.width),
            height: Math.max(0, rect.height),
        };
    }
    return {
        top: 0,
        left: 0,
        width: document.documentElement.clientWidth,
        height: window.innerHeight,
    };
}

export interface CommitDetailPopupLayout {
    left: number;
    top: number;
    maxHeight: number;
    height: number;
}

export function computeAnchoredDialogPosition(
    anchor: PopupAnchor,
    width: number,
    height: number,
    bounds?: ViewportBounds,
    variant: AnchoredDialogPositionVariant = 'default',
): { left: number; top: number } {
    const viewport = bounds ?? getViewportBounds();
    const rawTop = variant === 'merge'
        ? (anchor.repoToolbar && anchor.y < 120 ? anchor.y + 4 : anchor.y - 150)
        : (anchor.repoToolbar && anchor.y < 120 ? anchor.y + 8 : anchor.y - 90);
    const maxTop = Math.max(POPUP_MARGIN, viewport.height - height - POPUP_MARGIN);
    const top = Math.min(Math.max(rawTop, POPUP_MARGIN), maxTop);
    const maxLeft = Math.max(POPUP_MARGIN, viewport.width - width - POPUP_MARGIN);
    const preferredLeft = anchor.x - width / 2;
    const left = Math.min(Math.max(preferredLeft, POPUP_MARGIN), maxLeft);
    return { left, top };
}

function computePopupTop(
    anchorY: number,
    popupHeight: number,
    bodyTop: number,
    bodyBottom: number,
): number {
    const minTop = bodyTop + POPUP_MARGIN;
    const maxBottom = bodyBottom - POPUP_MARGIN;
    const maxTop = Math.max(minTop, maxBottom - popupHeight);
    const preferredTop = anchorY - ANCHOR_TOP_OFFSET;

    if (preferredTop >= minTop && preferredTop + popupHeight <= maxBottom) {
        return preferredTop;
    }

    let top = preferredTop;
    if (top + popupHeight > maxBottom) {
        top = maxBottom - popupHeight;
    }
    if (top < minTop) {
        top = minTop;
    }
    return Math.min(Math.max(top, minTop), maxTop);
}

export function computeCommitDetailPopupPosition(
    anchor: PopupAnchor,
    popupHeight: number,
    popupWidth = COMMIT_DETAIL_POPUP_WIDTH,
    bounds?: ViewportBounds,
): CommitDetailPopupLayout {
    const viewport = bounds ?? getViewportBounds();
    const bodyTop = viewport.top;
    const bodyBottom = viewport.top + viewport.height;
    const bodyRight = viewport.left + viewport.width;

    const maxHeight = Math.min(
        MAX_COMMIT_DETAIL_HEIGHT,
        Math.max(160, viewport.height - POPUP_MARGIN * 2),
    );
    const effectiveHeight = Math.min(popupHeight, maxHeight);

    const top = computePopupTop(anchor.y, effectiveHeight, bodyTop, bodyBottom);

    const preferredLeft = anchor.x + ANCHOR_LEFT_OFFSET;
    const minLeft = viewport.left + POPUP_MARGIN;
    const maxLeft = Math.max(minLeft, bodyRight - popupWidth - POPUP_MARGIN);
    const left = Math.min(Math.max(preferredLeft, minLeft), maxLeft);

    const maxHeightFromTop = Math.min(
        MAX_COMMIT_DETAIL_HEIGHT,
        Math.max(160, bodyBottom - top - POPUP_MARGIN),
    );
    const height = Math.min(effectiveHeight, maxHeightFromTop);

    return { left, top, maxHeight: maxHeightFromTop, height };
}

export function anchorFromMouseEvent(
    event: { clientX: number; clientY: number },
    repoToolbar = false,
    centerHorizontally = false,
): PopupAnchor {
    return { x: event.clientX, y: event.clientY, repoToolbar, centerHorizontally };
}

export function anchorFromElement(element: Element): PopupAnchor {
    const rect = element.getBoundingClientRect();
    return {
        x: rect.left + rect.width / 2,
        y: rect.top + rect.height / 2,
    };
}
