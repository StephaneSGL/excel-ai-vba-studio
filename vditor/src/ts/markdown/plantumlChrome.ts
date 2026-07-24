import {codicon} from "../util/codicon";

const PLANTUML_FIGURE_CLASS = "vditor-plantuml-figure";
const PLANTUML_CHROME_CLASS = "vditor-plantuml-chrome";

const openPlantumlUrl = (url: string, event: MouseEvent, link: HTMLAnchorElement, vditor?: IVditor) => {
    if (vditor && typeof vditor.options.onLinkClick === "function") {
        // 工具栏单击即打开：复用 onLinkClick，用 dblclick action 绕过 Ctrl/⌘ 门禁
        vditor.options.onLinkClick({
            type: "link",
            href: url,
            text: "PlantUML",
            element: link,
            action: "dblclick",
        }, event, vditor);
        return;
    }
    window.open(url, "_blank");
};

const createPlantumlChrome = (url: string, vditor?: IVditor) => {
    const chromeRoot = document.createElement("div");
    chromeRoot.className = PLANTUML_CHROME_CLASS;

    const toolbar = document.createElement("div");
    toolbar.className = "vditor-plantuml-chrome__toolbar";

    const spacer = document.createElement("div");
    spacer.className = "vditor-plantuml-chrome__spacer";
    toolbar.appendChild(spacer);

    const actions = document.createElement("div");
    actions.className = "vditor-plantuml-chrome__actions";

    const link = document.createElement("a");
    link.className = "vditor-plantuml-chrome__link";
    link.href = url;
    link.target = "_blank";
    link.rel = "noopener noreferrer";
    link.setAttribute("aria-label", "Open in PlantUML");
    link.innerHTML = `<span class="vditor-plantuml-chrome__link-icon">${codicon("link-external")}</span>`;
    actions.appendChild(link);

    toolbar.appendChild(actions);
    chromeRoot.appendChild(toolbar);

    link.addEventListener("mousedown", (event) => {
        event.preventDefault();
        event.stopPropagation();
    });
    link.addEventListener("click", (event) => {
        event.preventDefault();
        event.stopPropagation();
        openPlantumlUrl(url, event, link, vditor);
    });

    return chromeRoot;
};

/** 以图片为定位盒：figure 收缩到 img 尺寸，工具栏贴在图片右上角 */
export const ensurePlantumlChrome = (plantumlElement: HTMLElement, url: string, vditor?: IVditor) => {
    const img = plantumlElement.querySelector("img");
    if (!img) {
        return;
    }

    let figure = img.closest(`.${PLANTUML_FIGURE_CLASS}`) as HTMLElement | null;
    if (!figure) {
        figure = document.createElement("div");
        figure.className = PLANTUML_FIGURE_CLASS;
        img.replaceWith(figure);
        figure.appendChild(img);
    }

    figure.querySelector(`.${PLANTUML_CHROME_CLASS}`)?.remove();
    figure.insertBefore(createPlantumlChrome(url, vditor), figure.firstChild);
};
