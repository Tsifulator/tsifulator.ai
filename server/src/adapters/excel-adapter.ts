import { AdapterContext } from "../types";
import { AppAdapter } from "./contract";

export const excelAdapter: AppAdapter = {
  name: "excel",
  captureContext(context: AdapterContext) {
    return {
      mode: "excel-mock",
      note: "Excel adapter scaffold only in Phase 1",
      sessionId: context.sessionId
    };
  },
  proposeActions() {
    return {
      text: "Excel adapter scaffold is active. Deep Office execution is deferred for this phase."
    };
  },
  validateAction() {
    return "confirm";
  },
  async executeAction() {
    return {
      supported: false,
      output: "Excel execution not implemented in Phase 1"
    };
  },
  emitEvents(context: AdapterContext) {
    return [
      {
        type: "excel_adapter_scaffold_used",
        payload: { sessionId: context.sessionId }
      }
    ];
  }
};
