import React from "react";
import DashboardSuppliers from "./components/DashboardSuppliers";

type ErrorBoundaryState = { hasError: boolean; message: string };

class ErrorBoundary extends React.Component<
  { children: React.ReactNode },
  ErrorBoundaryState
> {
  constructor(props: { children: React.ReactNode }) {
    super(props);
    this.state = { hasError: false, message: "" };
  }

  static getDerivedStateFromError(error: unknown): ErrorBoundaryState {
    const message =
      error instanceof Error ? error.message : "Error desconocido";
    return { hasError: true, message };
  }

  handleReset = () => {
    this.setState({ hasError: false, message: "" });
  };

  render() {
    if (this.state.hasError) {
      return (
        <div className="flex min-h-screen flex-col items-center justify-center gap-6 bg-slate-50 px-4 text-center">
          <div className="rounded-2xl border border-red-200 bg-white px-8 py-10 shadow-lg max-w-md w-full">
            <p className="text-4xl mb-4">⚠️</p>
            <h1 className="text-lg font-bold text-slate-800 mb-2">
              Algo salió mal
            </h1>
            <p className="text-sm text-slate-500 mb-6 break-words">
              {this.state.message}
            </p>
            <button
              type="button"
              onClick={this.handleReset}
              className="rounded-xl bg-slate-900 px-6 py-2.5 text-sm font-semibold text-white hover:bg-slate-700"
            >
              Reintentar
            </button>
          </div>
        </div>
      );
    }
    return this.props.children;
  }
}

export default function App() {
  return (
    <ErrorBoundary>
      <DashboardSuppliers />
    </ErrorBoundary>
  );
}
