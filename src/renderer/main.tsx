import React, { Component, type ErrorInfo, type ReactNode } from "react";
import ReactDOM from "react-dom/client";
import App from "./App";
import "./styles.css";

class RendererErrorBoundary extends Component<{ children: ReactNode }, { error: string }> {
  state = { error: "" };

  static getDerivedStateFromError(error: unknown) {
    return { error: error instanceof Error ? error.message : String(error) };
  }

  componentDidCatch(error: Error, info: ErrorInfo) {
    void window.roadmapApi?.reportRendererDiagnostic({ event: "react-error", detail: `${error.name}: ${error.message}; component=${info.componentStack?.slice(0, 800) ?? "unknown"}` });
  }

  render() {
    if (this.state.error) return <div className="boot boot-error"><strong>页面初始化失败</strong><span>{this.state.error}</span><small>请从应用左侧或启动错误提示中打开诊断日志目录。</small></div>;
    return this.props.children;
  }
}

ReactDOM.createRoot(document.getElementById("root")!).render(
  <React.StrictMode>
    <RendererErrorBoundary><App /></RendererErrorBoundary>
  </React.StrictMode>,
);
