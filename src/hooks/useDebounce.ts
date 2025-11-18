import { useEffect, useState } from 'react';

/**
 * Custom hook to debounce a value
 * 
 * This hook delays updating the debounced value until after the specified delay
 * has elapsed since the last time the input value changed. This is useful for
 * preventing excessive API calls while the user is typing.
 * 
 * @param value - The value to debounce
 * @param delay - The delay in milliseconds (default: 400ms)
 * @returns The debounced value
 */
export function useDebounce<T>(value: T, delay: number = 400): T {
  const [debouncedValue, setDebouncedValue] = useState<T>(value);

  useEffect(() => {
    // Set up a timer to update the debounced value after the delay
    const handler = setTimeout(() => {
      setDebouncedValue(value);
    }, delay);

    // Clean up the timer if value changes before delay expires
    // This ensures we only call the API after user stops typing
    return () => {
      clearTimeout(handler);
    };
  }, [value, delay]);

  return debouncedValue;
}
