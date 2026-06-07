import { invoke } from "@tauri-apps/api/core";

export interface BusinessDictionaryItem {
  id: string;
  dictionaryId: string;
  value: string;
  label: string;
  description?: string | null;
  enabled: boolean;
  isDefault: boolean;
  sortOrder: number;
  createdAt: string;
  updatedAt: string;
}

export interface BusinessDictionary {
  id: string;
  scope: "workspace" | "user";
  dictionaryKey: string;
  name: string;
  description?: string | null;
  applicableFieldKeys: string[];
  enabled: boolean;
  sortOrder: number;
  createdAt: string;
  updatedAt: string;
  items: BusinessDictionaryItem[];
}

export interface BusinessDictionaryItemInput {
  id?: string | null;
  dictionaryId: string;
  value: string;
  label: string;
  description?: string | null;
  enabled?: boolean;
  isDefault?: boolean;
  sortOrder?: number;
}

export const businessDictionaryService = {
  list(includeDisabledItems = true): Promise<BusinessDictionary[]> {
    return invoke<BusinessDictionary[]>("list_business_dictionaries", {
      includeDisabledItems,
    });
  },

  getOptions(dictionaryKey: string): Promise<BusinessDictionaryItem[]> {
    return invoke<BusinessDictionaryItem[]>("get_business_dictionary_options", {
      dictionaryKey,
    });
  },

  saveItem(item: BusinessDictionaryItemInput): Promise<BusinessDictionaryItem> {
    return invoke<BusinessDictionaryItem>("save_business_dictionary_item", { item });
  },

  setItemEnabled(id: string, enabled: boolean): Promise<BusinessDictionaryItem> {
    return invoke<BusinessDictionaryItem>("set_business_dictionary_item_enabled", {
      id,
      enabled,
    });
  },

  deleteItem(id: string): Promise<void> {
    return invoke<void>("delete_business_dictionary_item", { id });
  },

  reorderItems(dictionaryId: string, itemIds: string[]): Promise<BusinessDictionaryItem[]> {
    return invoke<BusinessDictionaryItem[]>("reorder_business_dictionary_items", {
      dictionaryId,
      itemIds,
    });
  },
};
