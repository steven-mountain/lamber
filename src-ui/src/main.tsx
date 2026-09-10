import React from "react"
import ReactDOM from "react-dom/client"
import App from "./App"
import "./index.css"
import { useAppearanceStore } from "./store/useAppearanceStore"

// Retired runtime credentials must not remain in browser storage after upgrade.
// This is deletion only; dsh settings and chat history keep their own stores.
for (const key of ['lamber_ai_endpoint', 'lamber_ai_model', 'lamber_ai_api_key', 'lamber_ai_vision_enabled']) {
  localStorage.removeItem(key);
}

// Hydrate appearance settings synchronously early to prevent flashing on startup
useAppearanceStore.getState().hydrate();

ReactDOM.createRoot(document.getElementById("root") as HTMLElement).render(
  <React.StrictMode>
    <App />
  </React.StrictMode>
)
