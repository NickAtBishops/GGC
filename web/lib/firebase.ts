// Client-side Firebase initialization (modular SDK). No firebase-admin here —
// the engine verifies ID tokens server-side; this app only signs in, reads its
// own Firestore run history, and fetches finished models from Storage.
//
// Initialization is LAZY (first call wins, then cached): `getAuth()` validates
// the API key immediately, so eager module-level init would crash `next build`
// prerendering when env vars aren't present. All call sites run in effects or
// event handlers, i.e. strictly in the browser.

import { getApp, getApps, initializeApp, type FirebaseApp } from "firebase/app";
import { getAuth, GoogleAuthProvider, type Auth } from "firebase/auth";
import { getFirestore, type Firestore } from "firebase/firestore";
import { getStorage, type FirebaseStorage } from "firebase/storage";

function firebaseApp(): FirebaseApp {
  if (getApps().length > 0) return getApp();
  return initializeApp({
    apiKey: process.env.NEXT_PUBLIC_FIREBASE_API_KEY,
    authDomain: process.env.NEXT_PUBLIC_FIREBASE_AUTH_DOMAIN,
    projectId: process.env.NEXT_PUBLIC_FIREBASE_PROJECT_ID,
    storageBucket: process.env.NEXT_PUBLIC_FIREBASE_STORAGE_BUCKET,
    messagingSenderId: process.env.NEXT_PUBLIC_FIREBASE_MESSAGING_SENDER_ID,
    appId: process.env.NEXT_PUBLIC_FIREBASE_APP_ID,
  });
}

let cachedAuth: Auth | null = null;
export function firebaseAuth(): Auth {
  if (!cachedAuth) cachedAuth = getAuth(firebaseApp());
  return cachedAuth;
}

let cachedProvider: GoogleAuthProvider | null = null;
export function googleProvider(): GoogleAuthProvider {
  if (!cachedProvider) cachedProvider = new GoogleAuthProvider();
  return cachedProvider;
}

let cachedDb: Firestore | null = null;
export function firestoreDb(): Firestore {
  if (!cachedDb) cachedDb = getFirestore(firebaseApp());
  return cachedDb;
}

let cachedStorage: FirebaseStorage | null = null;
export function firebaseStorage(): FirebaseStorage {
  if (!cachedStorage) cachedStorage = getStorage(firebaseApp());
  return cachedStorage;
}
