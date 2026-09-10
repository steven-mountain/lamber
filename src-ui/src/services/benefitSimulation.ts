import { invoke } from '@tauri-apps/api/core';
import { listen } from '@tauri-apps/api/event';
import { prepareBenefitSimulation, type BenefitOverride } from '../lib/ictBenefitSimulation';
import { isTaxInclAutoFixEnabled } from '../store/useCalcPreferencesStore';
import type { IctInput } from '../utils/projectService';
/** Opaque jobs are claimed from Rust; models cannot supply complete inputs. No app/editor store writes. */
export async function listenBenefitSimulations() {
    return listen<string>('lamber-prepare-benefit-simulation', async (event) => {
        const id = event.payload;
        let prepared = null, error = null;
        try {
            const job = await invoke<{
                input: IctInput;
                overrides: BenefitOverride[];
            }>('ai_claim_benefit_simulation', { id });
            prepared = prepareBenefitSimulation(job.input, job.overrides, isTaxInclAutoFixEnabled());
        }
        catch (e) {
            error = e instanceof Error ? e.message : String(e);
        }
        try {
            await invoke('ai_finish_benefit_simulation', { id, prepared, error });
        }
        catch (e) {
            console.warn('只读试算请求已失效', e);
        }
    });
}
