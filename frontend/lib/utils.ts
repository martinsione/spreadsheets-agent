import { type ClassValue, clsx } from "clsx";
import { useCallback, useRef, useSyncExternalStore } from "react";
import { twMerge } from "tailwind-merge";

export function cn(...inputs: ClassValue[]) {
  return twMerge(clsx(inputs));
}

export function useLocalStorage<T>(
  key: string,
  initialValue: T,
): [T, (value: T | ((prev: T) => T)) => void] {
  const initialValueRef = useRef(initialValue);

  const getSnapshot = useCallback(() => {
    try {
      const item = localStorage.getItem(key);
      return item ? (JSON.parse(item) as T) : initialValueRef.current;
    } catch {
      return initialValueRef.current;
    }
  }, [key]);

  const getServerSnapshot = useCallback(() => initialValueRef.current, []);

  const subscribe = useCallback(
    (onStoreChange: () => void) => {
      const handleStorageChange = (e: StorageEvent) => {
        if (e.key === key) {
          onStoreChange();
        }
      };
      window.addEventListener("storage", handleStorageChange);
      return () => window.removeEventListener("storage", handleStorageChange);
    },
    [key],
  );

  const value = useSyncExternalStore(subscribe, getSnapshot, getServerSnapshot);

  const setValue = useCallback(
    (newValue: T | ((prev: T) => T)) => {
      const valueToStore =
        newValue instanceof Function ? newValue(getSnapshot()) : newValue;
      localStorage.setItem(key, JSON.stringify(valueToStore));
      window.dispatchEvent(
        new StorageEvent("storage", {
          key,
          newValue: JSON.stringify(valueToStore),
        }),
      );
    },
    [key, getSnapshot],
  );

  return [value, setValue];
}
