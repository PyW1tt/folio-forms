import { ArrowLeftRight } from "lucide-react";
import { useEffect, useState } from "react";

import { Button, Spinner } from "@/components/ui";
import {
  legacySsoEnabled,
  legacySsoReturnPath,
  startLegacySso,
} from "@/lib/api";
import { requestEditorSave } from "@/lib/form-lifecycle";

interface LegacySsoSwitchButtonProps {
  compact?: boolean;
  onError?: () => void;
  onStart?: () => void;
}

export const LegacySsoSwitchButton = ({
  compact = false,
  onError,
  onStart,
}: LegacySsoSwitchButtonProps) => {
  const [available, setAvailable] = useState(false);
  const [switching, setSwitching] = useState(false);
  const [error, setError] = useState(false);

  useEffect(() => {
    let active = true;
    const checkLegacySso = async () => {
      try {
        const enabled = await legacySsoEnabled();
        if (active) {
          setAvailable(enabled);
        }
      } catch {
        if (active) {
          setAvailable(false);
        }
      }
    };
    void checkLegacySso();
    return () => {
      active = false;
    };
  }, []);

  const switchAccount = async () => {
    if (switching) {
      return;
    }
    setSwitching(true);
    setError(false);
    try {
      if (!(await requestEditorSave())) {
        return;
      }
      const returnTo = legacySsoReturnPath(
        window.location.pathname,
        window.location.search
      );
      const { authorizationUrl } = await startLegacySso(returnTo);
      window.location.assign(authorizationUrl);
    } catch {
      setError(true);
      onError?.();
    } finally {
      setSwitching(false);
    }
  };

  if (!available) {
    return null;
  }

  return (
    <div className="flex flex-col items-end gap-1">
      <Button
        aria-label={compact ? "สลับบัญชีระบบเดิม" : undefined}
        disabled={switching}
        onClick={() => {
          onStart?.();
          void switchAccount();
        }}
        size="sm"
        type="button"
        variant="ghost"
      >
        {switching ? <Spinner /> : <ArrowLeftRight size={15} />}
        {compact ? "สลับบัญชี" : "สลับบัญชีระบบเดิม"}
      </Button>
      {error && !onError ? (
        <span className="text-xs text-[var(--danger)]" role="alert">
          ไม่สามารถเชื่อมต่อบัญชีระบบเดิมได้ กรุณาลองอีกครั้ง
        </span>
      ) : null}
    </div>
  );
};
