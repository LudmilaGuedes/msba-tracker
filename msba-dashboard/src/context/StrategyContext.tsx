import { createContext, useCallback, useContext, useEffect, useMemo, useState, type ReactNode } from "react";
import {
  loadStrategyData,
  loadStrategySettings,
  saveStrategyData,
  saveStrategySettings,
} from "../utils/storage";
import { DEFAULT_STRATEGY_SETTINGS } from "../types/strategy";
import type { ProgramStrategyData, StrategyDataStore, StrategySettings } from "../types/strategy";

interface StrategyContextValue {
  strategyData: StrategyDataStore;
  settings: StrategySettings;
  getStrategy: (programId: string) => ProgramStrategyData;
  updateStrategy: (programId: string, patch: Partial<ProgramStrategyData>) => void;
  updateSettings: (patch: Partial<StrategySettings>) => void;
}

const StrategyContext = createContext<StrategyContextValue | null>(null);

const EMPTY_STRATEGY: ProgramStrategyData = {};

export function StrategyProvider({ children }: { children: ReactNode }) {
  const [strategyData, setStrategyData] = useState<StrategyDataStore>({});
  const [settings, setSettings] = useState<StrategySettings>(DEFAULT_STRATEGY_SETTINGS);
  const [loaded, setLoaded] = useState(false);

  useEffect(() => {
    setStrategyData(loadStrategyData());
    setSettings(loadStrategySettings());
    setLoaded(true);
  }, []);

  useEffect(() => {
    if (loaded) saveStrategyData(strategyData);
  }, [strategyData, loaded]);

  useEffect(() => {
    if (loaded) saveStrategySettings(settings);
  }, [settings, loaded]);

  const getStrategy = useCallback(
    (programId: string) => strategyData[programId] ?? EMPTY_STRATEGY,
    [strategyData],
  );

  const updateStrategy = useCallback((programId: string, patch: Partial<ProgramStrategyData>) => {
    setStrategyData((prev) => ({
      ...prev,
      [programId]: { ...prev[programId], ...patch, updatedAt: Date.now() },
    }));
  }, []);

  const updateSettings = useCallback((patch: Partial<StrategySettings>) => {
    setSettings((prev) => ({ ...prev, ...patch }));
  }, []);

  const value = useMemo<StrategyContextValue>(
    () => ({ strategyData, settings, getStrategy, updateStrategy, updateSettings }),
    [strategyData, settings, getStrategy, updateStrategy, updateSettings],
  );

  return <StrategyContext.Provider value={value}>{children}</StrategyContext.Provider>;
}

export function useStrategy(): StrategyContextValue {
  const ctx = useContext(StrategyContext);
  if (!ctx) throw new Error("useStrategy deve ser usado dentro de um StrategyProvider");
  return ctx;
}
