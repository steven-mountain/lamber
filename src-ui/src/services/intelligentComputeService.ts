import { invoke } from "@tauri-apps/api/core";
import type {
  IntelligentAmountSource,
  IntelligentComputeProjectData,
  IntelligentComputeProjectState,
} from "../features/ai-compute-quote/types";

export const intelligentComputeService = {
  loadProject(projectId: string): Promise<IntelligentComputeProjectData> {
    return invoke<IntelligentComputeProjectData>("get_intelligent_compute_project", { projectId });
  },

  saveProjectState(
    projectId: string,
    request: {
      expectedVersion: number;
      activeAmountSourceId?: string | null;
      projectYears: number;
      discountRate: number;
    },
  ): Promise<IntelligentComputeProjectState> {
    return invoke<IntelligentComputeProjectState>("save_intelligent_compute_project_state", {
      projectId,
      request,
    });
  },

  saveAmountSource(
    projectId: string,
    source: IntelligentAmountSource,
    expectedVersion: number,
  ): Promise<IntelligentAmountSource> {
    return invoke<IntelligentAmountSource>("save_intelligent_amount_source", {
      projectId,
      request: { source, expectedVersion },
    });
  },

  deleteAmountSource(projectId: string, amountSourceId: string): Promise<void> {
    return invoke<void>("delete_intelligent_amount_source", { projectId, amountSourceId });
  },
};
