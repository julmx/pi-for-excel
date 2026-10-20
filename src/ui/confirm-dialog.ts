import {
  closeOverlayById,
  createOverlayButton,
  createOverlayDialog,
  createOverlayHeader,
} from "./overlay-dialog.js";
import { CONFIRM_DIALOG_OVERLAY_ID } from "./overlay-ids.js";
import { t } from "../language/index.js";

function getConfirmationUiUnavailableError(): string {
  return t("confirm.unavailable");
}

export type ConfirmButtonTone = "primary" | "danger";

export interface ConfirmDialogOptions {
  title: string;
  message: string;
  /** Optional full source code shown in an expandable section before approval. */
  codePreview?: string;
  confirmLabel?: string;
  cancelLabel?: string;
  confirmButtonTone?: ConfirmButtonTone;
  overlayId?: string;
  restoreFocusOnClose?: boolean;
  cardClassName?: string;
}

function canRenderConfirmationDialog(): boolean {
  if (typeof document === "undefined") {
    return false;
  }

  return document.body instanceof HTMLElement;
}

function getConfirmButtonClassName(tone: ConfirmButtonTone | undefined): string {
  return tone === "danger" ? "pi-overlay-btn--danger" : "pi-overlay-btn--primary";
}

export function requestConfirmationDialog(options: ConfirmDialogOptions): Promise<boolean> {
  if (!canRenderConfirmationDialog()) {
    return Promise.reject(new Error(getConfirmationUiUnavailableError()));
  }

  const overlayId = options.overlayId ?? CONFIRM_DIALOG_OVERLAY_ID;
  closeOverlayById(overlayId);

  return new Promise((resolve) => {
    const dialog = createOverlayDialog({
      overlayId,
      cardClassName: options.cardClassName ?? "pi-welcome-card pi-overlay-card pi-overlay-card--s",
      ...(options.restoreFocusOnClose !== undefined ? { restoreFocusOnClose: options.restoreFocusOnClose } : {}),
    });

    let settled = false;

    const settle = (confirmed: boolean): void => {
      if (settled) {
        return;
      }

      settled = true;
      resolve(confirmed);
    };

    const cancel = (): void => {
      settle(false);
      dialog.close();
    };

    const confirm = (): void => {
      settle(true);
      dialog.close();
    };

    const { header } = createOverlayHeader({
      onClose: cancel,
      closeLabel: options.cancelLabel ?? t("confirm.cancel"),
      title: options.title,
    });

    const body = document.createElement("div");
    body.className = "pi-overlay-body";

    const message = document.createElement("p");
    message.className = "pi-overlay-subtitle pi-confirm-dialog__message";
    message.textContent = options.message;

    // Let the user inspect the full Office.js script without expanding the
    // default confirmation dialog for every ordinary action.
    const codeDisclosure = document.createElement("details");
    if (options.codePreview !== undefined) {
      const summary = document.createElement("summary");
      summary.textContent = t("confirm.reviewFullCode");
      const code = document.createElement("pre");
      code.textContent = options.codePreview;
      code.style.maxHeight = "16rem";
      code.style.overflow = "auto";
      code.style.whiteSpace = "pre-wrap";
      code.style.overflowWrap = "anywhere";
      codeDisclosure.append(summary, code);
    }

    const actions = document.createElement("div");
    actions.className = "pi-overlay-actions";

    const cancelButton = createOverlayButton({
      text: options.cancelLabel ?? t("confirm.cancel"),
    });

    const confirmButton = createOverlayButton({
      text: options.confirmLabel ?? t("confirm.confirm"),
      className: getConfirmButtonClassName(options.confirmButtonTone),
    });

    cancelButton.addEventListener("click", cancel);
    confirmButton.addEventListener("click", confirm);

    dialog.addCleanup(() => {
      cancelButton.removeEventListener("click", cancel);
      confirmButton.removeEventListener("click", confirm);
      settle(false);
    });

    actions.append(cancelButton, confirmButton);
    body.appendChild(message);
    if (options.codePreview !== undefined) body.appendChild(codeDisclosure);
    dialog.card.append(header, body, actions);
    dialog.mount();
  });
}
