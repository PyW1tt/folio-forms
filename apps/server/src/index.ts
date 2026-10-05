import { ensureBootstrapAdmin } from "@onlyoffice/auth";

import { createApp } from "./app";
import { reconcileRecoverableState } from "./operations/recovery";

console.warn(
  "Single-host MMVP has no application or off-host backup; disk, ransomware, or regional loss is unrecoverable."
);
await ensureBootstrapAdmin();
await reconcileRecoverableState();

const recover = async (): Promise<void> => {
  try {
    await reconcileRecoverableState();
  } catch (error) {
    console.error("Recoverable state reconciliation failed", error);
  }
};
const recoveryInterval = setInterval(() => {
  void recover();
}, 60_000);
recoveryInterval.unref();
createApp().listen(3000, () => {
  console.log("Server is running on http://localhost:3000");
});
