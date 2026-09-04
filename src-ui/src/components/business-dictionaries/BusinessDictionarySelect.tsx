import { useEffect, useMemo, useState } from "react";
import { businessDictionaryService } from "../../services/businessDictionaryService";
import { useWorkspaceStore } from "../../store/useWorkspaceStore";

export interface DictionaryFallbackOption {
  value: string;
  label: string;
}

interface BusinessDictionarySelectProps {
  dictionaryKey: string;
  value: string;
  onChange: (value: string) => void;
  fallbackOptions: DictionaryFallbackOption[];
  name?: string;
  className?: string;
  disabled?: boolean;
  "aria-label"?: string;
}

export default function BusinessDictionarySelect({
  dictionaryKey,
  value,
  onChange,
  fallbackOptions,
  name,
  className = "",
  disabled = false,
  "aria-label": ariaLabel,
}: BusinessDictionarySelectProps) {
  const workspaceId = useWorkspaceStore(state => state.workspaceId);
  const [loadedOptions, setLoadedOptions] = useState<DictionaryFallbackOption[] | null>(null);

  useEffect(() => {
    let cancelled = false;
    setLoadedOptions(null);
    if (!workspaceId) return () => {
      cancelled = true;
    };

    void businessDictionaryService.getOptions(dictionaryKey)
      .then(items => {
        if (!cancelled) {
          setLoadedOptions(items.map(item => ({ value: item.value, label: item.label })));
        }
      })
      .catch(error => {
        console.warn(`Failed to load business dictionary "${dictionaryKey}"`, error);
        if (!cancelled) setLoadedOptions(null);
      });
    return () => {
      cancelled = true;
    };
  }, [dictionaryKey, workspaceId]);

  const options = useMemo(() => {
    const source = loadedOptions ?? fallbackOptions;
    const deduplicated = Array.from(new Map(source.map(item => [item.value, item])).values());
    const hasCurrentValue = !value || deduplicated.some(item => item.value === value);
    if (hasCurrentValue) return deduplicated;
    return [
      { value, label: `${value}（当前值，已停用或未配置）` },
      ...deduplicated,
    ];
  }, [fallbackOptions, loadedOptions, value]);

  return (
    <select
      name={name}
      value={value}
      disabled={disabled}
      aria-label={ariaLabel}
      onChange={event => onChange(event.target.value)}
      className={className}
    >
      {options.map(option => (
        <option key={option.value} value={option.value}>
          {option.label}
        </option>
      ))}
    </select>
  );
}
