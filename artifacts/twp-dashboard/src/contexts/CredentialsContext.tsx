import { createContext, useContext, useState, ReactNode } from "react";

interface Credentials {
  apiToken: string;
  phoneNumberId: string;
}

interface CredentialsContextType {
  credentials: Credentials | null;
  setCredentials: (creds: Credentials) => void;
  clearCredentials: () => void;
  isConfigured: boolean;
}

const CredentialsContext = createContext<CredentialsContextType | undefined>(undefined);

export function CredentialsProvider({ children }: { children: ReactNode }) {
  const [credentials, setCredentialsState] = useState<Credentials | null>(() => {
    try {
      const stored = localStorage.getItem("clarity_hub_credentials");
      if (stored) {
        const parsed = JSON.parse(stored);
        if (parsed.apiToken && parsed.phoneNumberId) return parsed;
      }
    } catch (_) {}
    return null;
  });

  const setCredentials = (creds: Credentials) => {
    localStorage.setItem("clarity_hub_credentials", JSON.stringify(creds));
    setCredentialsState(creds);
  };

  const clearCredentials = () => {
    localStorage.removeItem("clarity_hub_credentials");
    setCredentialsState(null);
  };

  return (
    <CredentialsContext.Provider
      value={{ credentials, setCredentials, clearCredentials, isConfigured: !!credentials?.apiToken && !!credentials?.phoneNumberId }}
    >
      {children}
    </CredentialsContext.Provider>
  );
}

export function useCredentials() {
  const ctx = useContext(CredentialsContext);
  if (!ctx) throw new Error("useCredentials must be used within CredentialsProvider");
  return ctx;
}
