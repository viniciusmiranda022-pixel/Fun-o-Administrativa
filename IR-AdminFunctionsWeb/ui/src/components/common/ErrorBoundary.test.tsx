import { describe, it, expect, vi } from 'vitest';
import { render, screen } from '@testing-library/react';
import ErrorBoundary from './ErrorBoundary';

function ThrowingComponent({ shouldThrow }: { shouldThrow: boolean }) {
  if (shouldThrow) throw new Error('Test error');
  return <div>OK</div>;
}

describe('ErrorBoundary', () => {
  it('renders children when no error', () => {
    render(
      <ErrorBoundary>
        <ThrowingComponent shouldThrow={false} />
      </ErrorBoundary>
    );
    expect(screen.getByText('OK')).toBeTruthy();
  });

  it('shows error UI when child throws', () => {
    // Suppress console.error for this test
    const consoleSpy = vi.spyOn(console, 'error').mockImplementation(() => {});
    render(
      <ErrorBoundary>
        <ThrowingComponent shouldThrow={true} />
      </ErrorBoundary>
    );
    expect(screen.getByText(/something went wrong/i)).toBeTruthy();
    consoleSpy.mockRestore();
  });
});
