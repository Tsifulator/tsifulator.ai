import { AdapterContext } from "../types";
import { AppAdapter } from "./contract";

export const rstudioAdapter: AppAdapter = {
  name: "rstudio",
  captureContext(context: AdapterContext) {
    return {
      mode: "rstudio-mock",
      note: "RStudio adapter scaffold — deep R integration deferred",
      sessionId: context.sessionId,
    };
  },
  proposeActions(context: AdapterContext) {
    const lower = context.message.toLowerCase();

    if (lower.startsWith("cmd:")) {
      return {
        text: "I drafted a command proposal from your explicit cmd request.",
        command: context.message.slice(4).trim(),
      };
    }

    if (lower.includes("run script") || lower.includes("source")) {
      return {
        text: "I can source an R script once you approve execution.",
        command: "Rscript --vanilla script.R",
      };
    }

    if (lower.includes("install package") || lower.includes("install.packages")) {
      return {
        text: "I can install an R package after your approval.",
        command: 'Rscript -e "install.packages(\'tidyverse\', repos=\'https://cran.r-project.org\')"',
      };
    }

    return {
      text: "RStudio adapter scaffold is active. Use 'cmd:' prefix for explicit commands, or ask to run/source scripts.",
    };
  },
  validateAction(command: string) {
    const lower = command.toLowerCase();
    // Block system-level commands from R context
    if (lower.includes("system(") || lower.includes("system2(")) {
      return "confirm";
    }
    if (lower.includes("unlink(") && lower.includes("recursive")) {
      return "blocked";
    }
    return "safe";
  },
  async executeAction() {
    return {
      supported: false,
      output: "RStudio execution not implemented — handled by core policy executor",
    };
  },
  emitEvents(context: AdapterContext) {
    return [
      {
        type: "rstudio_adapter_scaffold_used",
        payload: { sessionId: context.sessionId },
      },
    ];
  },
};
