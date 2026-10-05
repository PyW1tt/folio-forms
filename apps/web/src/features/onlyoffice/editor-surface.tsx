import { useEffect } from "react";
import type { RefObject } from "react";

import { Button, Notice, Spinner } from "@/components/ui";

import type { EditorConfig, OnlyOfficeEditorState } from "./editor-protocol";

export const EditorSurface = ({
  config,
  editorState,
  error,
  title,
  expanded,
  hostRef,
  surfaceRef,
  feedbackRef,
  surfaceId,
  keyboardHelpId,
  onRetry,
  onToggleExpanded,
  onRestore,
}: {
  config: EditorConfig | null;
  editorState: OnlyOfficeEditorState;
  error: string | null;
  title: string;
  expanded: boolean;
  hostRef: RefObject<HTMLDivElement | null>;
  surfaceRef: RefObject<HTMLDivElement | null>;
  feedbackRef: RefObject<HTMLDivElement | null>;
  surfaceId: string;
  keyboardHelpId: string;
  onRetry: () => void;
  onToggleExpanded: () => void;
  onRestore: () => void;
}) => {
  useEffect(() => {
    const surface = surfaceRef.current;
    if (!expanded || !surface) {
      return;
    }

    const obscured: HTMLElement[] = [];
    const focusableAncestors: [HTMLElement, string | null][] = [];
    const branches = new Map<HTMLElement, HTMLElement>();
    const hideSibling = (sibling: Element, activeBranch: HTMLElement) => {
      if (
        sibling instanceof HTMLElement &&
        sibling !== activeBranch &&
        !sibling.inert
      ) {
        sibling.inert = true;
        obscured.push(sibling);
      }
    };
    let branch: HTMLElement = surface;
    while (branch.parentElement) {
      const parent = branch.parentElement;
      branches.set(parent, branch);
      for (const sibling of parent.children) {
        hideSibling(sibling, branch);
      }
      branch = parent;
      if (parent.tabIndex >= 0) {
        focusableAncestors.push([parent, parent.getAttribute("tabindex")]);
        parent.tabIndex = -1;
      }
    }
    const observer = new MutationObserver((records) => {
      for (const record of records) {
        const activeBranch =
          record.target instanceof HTMLElement
            ? branches.get(record.target)
            : undefined;
        if (!activeBranch) {
          continue;
        }
        for (const sibling of record.addedNodes) {
          if (sibling instanceof Element) {
            hideSibling(sibling, activeBranch);
          }
        }
      }
    });
    for (const parent of branches.keys()) {
      observer.observe(parent, { childList: true });
    }

    const handleKeyDown = (event: KeyboardEvent) => {
      if (event.key === "Escape") {
        event.preventDefault();
        onRestore();
      }
    };

    window.addEventListener("keydown", handleKeyDown);
    return () => {
      window.removeEventListener("keydown", handleKeyDown);
      observer.disconnect();
      for (const sibling of obscured) {
        sibling.inert = false;
      }
      for (const [ancestor, tabIndex] of focusableAncestors) {
        if (tabIndex === null) {
          ancestor.removeAttribute("tabindex");
        } else {
          ancestor.setAttribute("tabindex", tabIndex);
        }
      }
    };
  }, [expanded, onRestore, surfaceRef]);

  const renderEditorContent = () => {
    const blocked = editorState === "blocked";
    if (blocked || editorState === "error" || error) {
      return (
        <div
          ref={feedbackRef}
          className={`grid ${
            expanded ? "h-full min-h-0" : "min-h-[520px]"
          } place-items-center p-8`}
          tabIndex={-1}
        >
          <Notice tone="danger">
            <div className="space-y-3">
              {blocked ? (
                <>
                  <p className="font-semibold">เอกสารนี้กำลังถูกแก้ไขโดยผู้ใช้รายอื่น</p>
                  <p>ยังไม่เปิดตัวแก้ไขจนกว่าจะเชื่อมต่อใหม่ได้</p>
                </>
              ) : (
                <p>{error ?? "ไม่สามารถเปิดตัวแก้ไขเอกสารได้"}</p>
              )}
              <Button type="button" variant="secondary" onClick={onRetry}>
                {blocked ? "ลองเชื่อมต่อใหม่" : "ลองใหม่"}
              </Button>
            </div>
          </Notice>
        </div>
      );
    }

    if (!config) {
      return (
        <div
          className={`grid ${
            expanded ? "h-full min-h-0" : "min-h-[520px]"
          } place-items-center gap-3 p-8 text-center`}
          aria-busy="true"
          role="status"
        >
          <Spinner />
          <p className="text-sm text-[var(--ink-soft)]">
            กำลังเตรียมตัวแก้ไขเอกสาร…
          </p>
        </div>
      );
    }

    if (config.editorUrl) {
      return (
        <iframe
          title={title}
          src={config.editorUrl}
          className={
            expanded
              ? "h-full min-h-0 w-full border-0"
              : "h-[min(72vh,760px)] min-h-[520px] w-full border-0"
          }
        />
      );
    }

    return (
      <div
        className={
          expanded
            ? "h-full min-h-0 w-full"
            : "h-[min(72vh,760px)] min-h-[520px] w-full"
        }
      >
        <div ref={hostRef} className="h-full w-full" aria-label={title} />
      </div>
    );
  };

  return (
    <div
      aria-label={expanded ? title : undefined}
      aria-describedby={expanded ? keyboardHelpId : undefined}
      aria-modal={expanded ? true : undefined}
      ref={surfaceRef}
      role={expanded ? "dialog" : undefined}
      className={
        expanded
          ? "fixed inset-0 z-[100] flex min-h-0 flex-col overflow-hidden overscroll-none bg-[var(--canvas)] p-3 sm:p-4"
          : "relative"
      }
    >
      <div className="flex shrink-0 flex-wrap items-center justify-end gap-2 pb-2">
        {expanded && (
          <p
            className="mr-auto text-sm text-[var(--ink-soft)]"
            id={keyboardHelpId}
          >
            คืนขนาดจากในเอกสารด้วยแป้นพิมพ์: กด Alt/Option แล้ว F และเลือก
            “คืนค่าขนาดปกติ” ตามคำใบ้ของ ONLYOFFICE
          </p>
        )}
        <Button
          aria-controls={surfaceId}
          aria-expanded={expanded}
          className="focus-visible:outline-2 focus-visible:outline-offset-2 focus-visible:outline-[var(--ink)]"
          onClick={onToggleExpanded}
          size="sm"
          type="button"
          variant="secondary"
        >
          {expanded ? "คืนค่าขนาดปกติ" : "ขยายพื้นที่เอกสาร"}
        </Button>
      </div>
      <div className={expanded ? "min-h-0 flex-1" : undefined} id={surfaceId}>
        {renderEditorContent()}
      </div>
    </div>
  );
};
