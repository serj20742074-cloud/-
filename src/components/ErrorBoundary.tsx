import React, { Component, ErrorInfo, ReactNode } from 'react';

interface Props {
  children: ReactNode;
}

interface State {
  hasError: boolean;
  error: Error | null;
}

export class ErrorBoundary extends Component<Props, State> {
  public state: State = {
    hasError: false,
    error: null
  };

  public static getDerivedStateFromError(error: Error): State {
    // Update state so the next render will show the fallback UI.
    return { hasError: true, error };
  }

  public componentDidCatch(error: Error, errorInfo: ErrorInfo) {
    console.error('Uncaught error inside React tree:', error, errorInfo);
  }

  private handleReload = () => {
    window.location.reload();
  };

  private handleResetAndReload = () => {
    // Redirect to the cache clearing URL to completely flush corrupt data/Service Workers
    const cleanUrl = window.location.protocol + '//' + window.location.host + window.location.pathname + '?clear-cache=true';
    window.location.replace(cleanUrl);
  };

  public render() {
    if (this.state.hasError) {
      return (
        <div className="min-h-screen bg-[#0f172a] text-[#f8fafc] flex flex-col items-center justify-center p-6 font-sans">
          <div className="max-w-2xl w-full bg-[#1e293b] rounded-2xl border border-slate-700 p-8 shadow-2xl space-y-6">
            <div className="flex items-center gap-4 text-rose-500 border-b border-slate-700/60 pb-4">
              <svg xmlns="http://www.w3.org/2000/svg" className="h-10 w-10 shrink-0" fill="none" viewBox="0 0 24 24" stroke="currentColor" strokeWidth={2}>
                <path strokeLinecap="round" strokeLinejoin="round" d="M12 9v2m0 4h.01m-6.938 4h13.856c1.54 0 2.502-1.667 1.732-3L13.732 4c-.77-1.333-2.694-1.333-3.464 0L3.34 16c-.77 1.333.192 3 1.732 3z" />
              </svg>
              <div>
                <h1 className="text-xl font-black tracking-tight uppercase">Произошла ошибка при отрисовке</h1>
                <p className="text-xs text-slate-400 font-medium">Обнаружено непредвиденное исключение в React-компоненте</p>
              </div>
            </div>

            <div className="space-y-3">
              <p className="text-sm text-slate-300">
                Приложение столкнулось с внутренней ошибкой. Это может быть связано с устаревшим кэшем Service Worker, поврежденными локальными данными в браузере или некорректным состоянием локальной базы данных.
              </p>
              
              {this.state.error && (
                <div className="p-4 bg-slate-900 rounded-xl border border-slate-800 font-mono text-xs text-rose-400 overflow-x-auto max-h-40 whitespace-pre-wrap">
                  <strong>{this.state.error.name}:</strong> {this.state.error.message}
                </div>
              )}
            </div>

            <div className="flex flex-wrap gap-3 border-t border-slate-700/60 pt-6">
              <button
                onClick={this.handleReload}
                className="bg-blue-600 hover:bg-blue-500 text-white font-bold text-sm px-6 py-3 rounded-xl transition flex items-center gap-2 shadow-lg shadow-blue-500/10 active:scale-[0.98]"
              >
                <svg xmlns="http://www.w3.org/2000/svg" className="h-4 w-4" fill="none" viewBox="0 0 24 24" stroke="currentColor" strokeWidth={2.5}>
                  <path strokeLinecap="round" strokeLinejoin="round" d="M4 4v5h.582m15.356 2A8.001 8.001 0 1121.21 15H19" />
                </svg>
                Перезагрузить страницу
              </button>
              
              <button
                onClick={this.handleResetAndReload}
                className="bg-rose-500/10 hover:bg-rose-500/20 text-rose-400 border border-rose-500/30 font-bold text-sm px-6 py-3 rounded-xl transition flex items-center gap-2 active:scale-[0.98]"
              >
                <svg xmlns="http://www.w3.org/2000/svg" className="h-4 w-4" fill="none" viewBox="0 0 24 24" stroke="currentColor" strokeWidth={2.5}>
                  <path strokeLinecap="round" strokeLinejoin="round" d="M19 7l-.867 12.142A2 2 0 0116.138 21H7.862a2 2 0 01-1.995-1.858L5 7m5 4v6m4-6v6m1-10V4a1 1 0 00-1-1h-4a1 1 0 00-1 1v3M4 7h16" />
                </svg>
                Очистить локальный кэш и базу
              </button>
            </div>
            
            <div className="text-[10px] text-slate-500 text-center uppercase tracking-wider">
              Информационно-аналитический комплекс учета энергоресурсов
            </div>
          </div>
        </div>
      );
    }

    return this.props.children;
  }
}
