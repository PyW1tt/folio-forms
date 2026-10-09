// oxlint-disable prefer-await-to-then -- Promise chaining serializes Office commands and their control snapshots.
import { applyNewDateDefaultsCommand } from "./document-commands.js";
import { parseCommandResult } from "./values.js";

export const createDateDefaults = (office, onChanged, onError) => {
  let knownIds = null;
  let started = false;
  let pending = Promise.resolve();
  const refresh = () => {
    if (!started) {
      return pending;
    }
    pending = pending
      .then(async () => {
        window.Asc.scope ||= {};
        const { scope } = window.Asc;
        scope.formBridgeKnownControlIds = knownIds;
        const result = parseCommandResult(
          await office.callCommandResult(applyNewDateDefaultsCommand)
        );
        knownIds = result.ids;
        if (result.changed) {
          onChanged();
        }
      })
      .catch(onError);
    return pending;
  };
  const start = () => {
    if (started) {
      return pending;
    }
    started = true;
    return refresh();
  };
  return { refresh, start };
};
