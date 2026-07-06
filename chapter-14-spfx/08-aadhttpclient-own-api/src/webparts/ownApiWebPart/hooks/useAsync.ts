import { useCallback, useEffect, useState } from 'react';

export interface IAsyncState<T> {
  data: T | undefined;
  loading: boolean;
  error: Error | undefined;
  refetch: () => void;
}

/** React hook for an async call (here a fetch to a custom API). */
export function useAsync<T>(
  fetcher: () => Promise<T>,
  deps: unknown[]
): IAsyncState<T> {
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
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, deps);

  useEffect((): void => {
    fetchData().catch(() => { /* handled inside fetchData */ });
  }, [fetchData]);

  return {
    data,
    loading,
    error,
    refetch: (): void => { fetchData().catch(() => { /* handled inside fetchData */ }); }
  };
}