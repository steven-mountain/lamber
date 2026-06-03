import React from "react"
import ReactDOM from "react-dom/client"
import App from "./App"
import "./index.css"
import { useAppearanceStore } from "./store/useAppearanceStore"

// Hydrate appearance settings synchronously early to prevent flashing on startup
useAppearanceStore.getState().hydrate();

ReactDOM.createRoot(document.getElementById("root") as HTMLElement).render(
  <React.StrictMode>
    <App />
  </React.StrictMode>
)
