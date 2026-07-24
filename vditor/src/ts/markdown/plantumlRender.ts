import {Constants} from "../constants";
import {addScript} from "../util/addScript";
import {plantumlRenderAdapter} from "./adapterRender";
import {ensurePlantumlChrome} from "./plantumlChrome";

declare const plantumlEncoder: {
    encode(options: string): string,
};

const PLANTUML_SOURCE_ATTR = "data-plantuml";

const buildPlantumlUrl = (text: string) =>
    `http://www.plantuml.com/plantuml/svg/~1${plantumlEncoder.encode(text)}`;

export const plantumlRender = (
    element: (HTMLElement | Document) = document,
    cdn = Constants.CDN,
    vditor?: IVditor,
) => {
    const plantumlElements = plantumlRenderAdapter.getElements(element);
    if (plantumlElements.length === 0) {
        return;
    }
    addScript(`${cdn}/dist/js/plantuml/plantuml-encoder.min.js`, "vditorPlantumlScript").then(() => {
        plantumlElements.forEach((e: HTMLDivElement) => {
            if (e.parentElement.classList.contains("vditor-wysiwyg__pre") ||
                e.parentElement.classList.contains("vditor-ir__marker--pre")) {
                return;
            }
            const text = plantumlRenderAdapter.getCode(e).trim()
                || e.getAttribute(PLANTUML_SOURCE_ATTR)?.trim()
                || "";
            if (!text) {
                return;
            }
            try {
                const url = buildPlantumlUrl(text);
                e.setAttribute(PLANTUML_SOURCE_ATTR, text);
                e.innerHTML = `<img src="${url}">`;
                ensurePlantumlChrome(e, url, vditor);
            } catch (error) {
                e.className = "vditor-reset--error";
                e.innerHTML = `plantuml render error: <br>${error}`;
            }
        });
    });
};
