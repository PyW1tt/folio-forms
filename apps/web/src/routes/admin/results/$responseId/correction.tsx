// oxlint-disable unicorn(filename-case) -- TanStack Router requires this dynamic route filename.
import {
  createFileRoute,
  Link,
  useNavigate,
  useParams,
} from "@tanstack/react-router";
import { useEffect, useRef, useState } from "react";

import { PageHeader } from "@/components/app-shell";
import { OnlyOfficeEditor } from "@/components/onlyoffice-editor";
import { Button, Card, Input, Notice } from "@/components/ui";
import type { EditorBridgeMessage } from "@/features/onlyoffice/editor-protocol";

const AdminCorrectionRoute = () => {
  const navigate = useNavigate();
  const { responseId } = useParams({
    from: "/admin/results/$responseId/correction",
  });
  const [reason, setReason] = useState("");
  const [saveRequest, setSaveRequest] = useState(0);
  const [message, setMessage] = useState<string | null>(null);
  const [saving, setSaving] = useState(false);
  const feedbackRef = useRef<HTMLDivElement>(null);
  useEffect(() => {
    if (message && !saving) {
      feedbackRef.current?.focus();
    }
  }, [message, saving]);

  const handleBridgeMessage = (bridgeMessage: EditorBridgeMessage) => {
    if (
      bridgeMessage.type !== "operation" ||
      bridgeMessage.action !== "save-correction"
    ) {
      return;
    }
    if (bridgeMessage.status === "pending") {
      setSaving(true);
      setMessage("กำลังบันทึก Correction…");
      return;
    }
    if (bridgeMessage.status === "failed") {
      setSaving(false);
      setMessage(bridgeMessage.error ?? "บันทึก Correction ไม่สำเร็จ กรุณาลองใหม่");
      return;
    }
    setSaving(false);
    void navigate({
      params: { responseId },
      search: { form: undefined },
      to: "/admin/results/$responseId",
    });
  };

  const submitCorrection = () => {
    const trimmedReason = reason.trim();
    if (!trimmedReason || saving) {
      setMessage("กรุณาระบุเหตุผลการแก้ไข");
      return;
    }
    setMessage(null);
    setSaveRequest((request) => request + 1);
  };

  return (
    <div className="space-y-6">
      <PageHeader
        title="แก้ไขคำตอบที่ส่งแล้ว"
        description="Correction จะสร้างฉบับใหม่โดยไม่แก้ไข Submission เดิม"
        action={
          <Link
            search={{ form: undefined }}
            to="/admin/results/$responseId"
            params={{ responseId }}
          >
            <Button variant="secondary">ยกเลิก</Button>
          </Link>
        }
      />
      <Card className="space-y-4 p-5">
        <label className="grid gap-1 text-sm font-semibold text-[var(--ink-soft)]">
          เหตุผลการแก้ไข
          <Input
            aria-describedby="correction-reason-help"
            aria-label="เหตุผลการแก้ไข"
            maxLength={2000}
            onChange={(event) => setReason(event.target.value)}
            placeholder="ระบุเหตุผลที่ตรวจสอบได้"
            required
            value={reason}
          />
        </label>
        <p
          id="correction-reason-help"
          className="text-sm text-[var(--ink-soft)]"
        >
          ระบบจะบันทึกผู้แก้ไข เวลา เหตุผล และข้อมูลฉบับใหม่เป็นประวัติถาวร
        </p>
        <div ref={feedbackRef} tabIndex={-1} className="outline-none">
          {message ? (
            <Notice tone={saving ? "neutral" : "danger"}>{message}</Notice>
          ) : null}
        </div>
        <div>
          <Button
            disabled={saving || reason.trim().length === 0}
            onClick={submitCorrection}
          >
            บันทึก Correction
          </Button>
        </div>
      </Card>
      <Card className="min-h-[640px] overflow-hidden p-0">
        <OnlyOfficeEditor
          configUrl={`/api/admin/results/${responseId}/correction/editor-config`}
          onBridgeMessage={handleBridgeMessage}
          saveAction="save-correction"
          saveReason={reason}
          saveRequest={saveRequest}
          title="Correction editor"
        />
      </Card>
    </div>
  );
};

export const Route = createFileRoute("/admin/results/$responseId/correction")({
  component: AdminCorrectionRoute,
});
