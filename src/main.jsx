import React from "react";
import ReactDOM from "react-dom/client";
import App from "./App.jsx";
import "./index.css";

// Shows a message instead of a blank page if the app ever crashes.
class ErrorBoundary extends React.Component {
  constructor(props) { super(props); this.state = { error: null }; }
  static getDerivedStateFromError(error) { return { error }; }
  render() {
    if (!this.state.error) return this.props.children;
    return (
      <div style={{ fontFamily: "Jost, system-ui, Arial, sans-serif", padding: "32px 20px", maxWidth: 520, margin: "0 auto", color: "#231f20" }}>
        <h1 style={{ fontWeight: 300, letterSpacing: ".14em", textTransform: "uppercase", color: "#6e6b6c", fontSize: 20, margin: "0 0 12px" }}>EcoOils Security</h1>
        <p style={{ fontSize: 16 }}>Something went wrong. Your saved data is safe. Reload to try again.</p>
        <p style={{ fontSize: 12, color: "#6e6b6c", wordBreak: "break-word" }}>{String(this.state.error && this.state.error.message || this.state.error)}</p>
        <button onClick={() => window.location.reload()} style={{ background: "#44a276", color: "#231f20", fontWeight: 700, border: 0, borderRadius: 8, padding: "12px 24px", fontSize: 16, cursor: "pointer" }}>Reload</button>
      </div>
    );
  }
}

ReactDOM.createRoot(document.getElementById("root")).render(
  <React.StrictMode>
    <ErrorBoundary>
      <App />
    </ErrorBoundary>
  </React.StrictMode>
);
