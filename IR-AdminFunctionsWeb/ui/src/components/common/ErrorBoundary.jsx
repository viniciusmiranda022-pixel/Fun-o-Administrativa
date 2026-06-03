import { Component } from 'react';

export default class ErrorBoundary extends Component {
  constructor(props) {
    super(props);
    this.state = { hasError: false, error: null };
  }

  static getDerivedStateFromError(error) {
    return { hasError: true, error };
  }

  componentDidCatch(error, info) {
    console.error('ErrorBoundary caught:', error, info);
  }

  render() {
    if (this.state.hasError) {
      return (
        <div className="flex flex-col items-center justify-center h-full p-8 text-center">
          <div className="text-4xl mb-4" style={{ color: '#DC3545' }}>⚠</div>
          <h2 className="text-base font-semibold text-[#222] mb-2">Something went wrong</h2>
          <p className="text-xs text-[#666] mb-4 max-w-md">
            {this.state.error?.message || 'An unexpected error occurred in this section.'}
          </p>
          <button
            onClick={() => this.setState({ hasError: false, error: null })}
            className="px-4 py-1.5 text-xs bg-[#0078A8] text-white hover:bg-[#006090]"
          >
            Try again
          </button>
        </div>
      );
    }
    return this.props.children;
  }
}
