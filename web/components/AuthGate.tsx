"use client";

// Gates every page behind Firebase Auth (Google sign-in). Unauthenticated
// visitors see a sign-in screen; signed-in users get the app plus a context
// providing the current User to all child components.

import { FirebaseError } from "firebase/app";
import { onAuthStateChanged, signInWithPopup, type User } from "firebase/auth";
import {
  createContext,
  useContext,
  useEffect,
  useState,
  type ReactNode,
} from "react";
import { firebaseAuth, googleProvider } from "@/lib/firebase";

const AuthContext = createContext<User | null>(null);

/** Current signed-in user. Only valid inside <AuthGate> (i.e. anywhere in the app). */
export function useUser(): User {
  const user = useContext(AuthContext);
  if (!user) {
    throw new Error("useUser() must be called from a component rendered inside <AuthGate>.");
  }
  return user;
}

function BuildingLogo() {
  return (
    <div className="w-14 h-14 rounded-2xl bg-linear-to-br from-blue-500 to-blue-700 flex items-center justify-center shadow-lg">
      <svg className="w-8 h-8 text-white" fill="none" stroke="currentColor" viewBox="0 0 24 24">
        <path
          strokeLinecap="round"
          strokeLinejoin="round"
          strokeWidth={2}
          d="M19 21V5a2 2 0 00-2-2H7a2 2 0 00-2 2v16m14 0h2m-2 0h-5m-9 0H3m2 0h5M9 7h1m-1 4h1m4-4h1m-1 4h1m-5 10v-5a1 1 0 011-1h2a1 1 0 011 1v5m-4 0h4"
        />
      </svg>
    </div>
  );
}

function GoogleG() {
  return (
    <svg className="w-5 h-5" viewBox="0 0 48 48" aria-hidden="true">
      <path
        fill="#FFC107"
        d="M43.611 20.083H42V20H24v8h11.303c-1.649 4.657-6.08 8-11.303 8-6.627 0-12-5.373-12-12s5.373-12 12-12c3.059 0 5.842 1.154 7.961 3.039l5.657-5.657C34.046 6.053 29.268 4 24 4 12.955 4 4 12.955 4 24s8.955 20 20 20 20-8.955 20-20c0-1.341-.138-2.65-.389-3.917z"
      />
      <path
        fill="#FF3D00"
        d="M6.306 14.691l6.571 4.819C14.655 15.108 18.961 12 24 12c3.059 0 5.842 1.154 7.961 3.039l5.657-5.657C34.046 6.053 29.268 4 24 4 16.318 4 9.656 8.337 6.306 14.691z"
      />
      <path
        fill="#4CAF50"
        d="M24 44c5.166 0 9.86-1.977 13.409-5.192l-6.19-5.238A11.91 11.91 0 0124 36c-5.202 0-9.619-3.317-11.283-7.946l-6.522 5.025C9.505 39.556 16.227 44 24 44z"
      />
      <path
        fill="#1976D2"
        d="M43.611 20.083H42V20H24v8h11.303a12.04 12.04 0 01-4.087 5.571l.003-.002 6.19 5.238C36.971 39.205 44 34 44 24c0-1.341-.138-2.65-.389-3.917z"
      />
    </svg>
  );
}

export default function AuthGate({ children }: { children: ReactNode }) {
  const [user, setUser] = useState<User | null>(null);
  const [ready, setReady] = useState(false);
  const [signingIn, setSigningIn] = useState(false);
  const [signInError, setSignInError] = useState<string | null>(null);

  useEffect(() => {
    const unsubscribe = onAuthStateChanged(firebaseAuth(), (u) => {
      setUser(u);
      setReady(true);
    });
    return unsubscribe;
  }, []);

  async function handleSignIn() {
    setSigningIn(true);
    setSignInError(null);
    try {
      await signInWithPopup(firebaseAuth(), googleProvider());
    } catch (e) {
      if (
        e instanceof FirebaseError &&
        (e.code === "auth/popup-closed-by-user" || e.code === "auth/cancelled-popup-request")
      ) {
        // User dismissed the popup — not an error worth surfacing.
      } else {
        setSignInError(e instanceof Error ? e.message : "Sign-in failed. Try again.");
      }
    } finally {
      setSigningIn(false);
    }
  }

  if (!ready) {
    return (
      <div className="min-h-screen flex items-center justify-center">
        <div className="flex items-center gap-3 text-slate-400">
          <div className="pulse-dot w-3 h-3 rounded-full bg-blue-400" />
          <span className="text-sm">Loading…</span>
        </div>
      </div>
    );
  }

  if (!user) {
    return (
      <div className="min-h-screen flex items-center justify-center p-6">
        <div className="card rounded-2xl p-8 w-full max-w-md text-center">
          <div className="flex justify-center mb-5">
            <BuildingLogo />
          </div>
          <h1 className="text-2xl font-bold mb-1">GGC Deal Engine</h1>
          <p className="text-slate-400 text-sm mb-8">
            AI underwriting for manufactured-housing communities and RV parks. Sign in with your
            authorized Google account to continue.
          </p>
          <button
            onClick={() => void handleSignIn()}
            disabled={signingIn}
            className="w-full flex items-center justify-center gap-3 bg-white text-slate-900 font-semibold rounded-lg py-3 hover:bg-slate-100 transition-colors disabled:opacity-50"
          >
            <GoogleG />
            {signingIn ? "Signing in…" : "Sign in with Google"}
          </button>
          {signInError && <p className="text-sm text-red-400 mt-4">{signInError}</p>}
          <p className="text-xs text-slate-500 mt-6">
            Access is restricted to allowlisted Gary Group Capital users.
          </p>
        </div>
      </div>
    );
  }

  return <AuthContext.Provider value={user}>{children}</AuthContext.Provider>;
}
