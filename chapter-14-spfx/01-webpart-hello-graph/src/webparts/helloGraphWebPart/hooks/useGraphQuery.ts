import { useCallback, useEffect, useState } from 'react';

export interface IGraphQueryState<T> {
  data: T | undefined;
  loading: boolean;
  error: Error | undefined;
  refetch: () => void;
}

/**
 * React hook that runs an async fetcher (e.g. a Graph query) and tracks
 * loading / data / error. Re-runs whenever any value in `deps` changes.
 *
 * Mirrors the hook shown in the book (chapter 14, "React Hooks with Graph").
 */
export function useGraphQuery<T>(
  fetcher: () => Promise<T>,
  deps: unknown[]
): IGraphQueryState<T> {
  const [data, setData] = useState<T | undefined>(undefined);
  const [loading, setLoading] = useState<boolean>(true);
  const [error, setError] = useState<Error | undefined>(undefined);

  const fetchData = useCallback(async (): Promise<void> => {
    setLoading(true);
    setError(undefined);
    try {
      const result: T = await fetcher();
      setData(result);
    } catch (err) {
      setError(err as Error);
    } finally {
      setLoading(false);
    }
    // fetcher is rebuilt by the caller when `deps` change, so we key on it.
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, deps);

  useEffect((): void => {
    // fetchData swallows its own errors into `error` state; .catch satisfies
    // no-floating-promises without re-throwing.
    fetchData().catch(() => { /* handled inside fetchData */ });
  }, [fetchData]);

  return {
    data,
    loading,
    error,
    refetch: (): void => { fetchData().catch(() => { /* handled inside fetchData */ }); }
  };
}