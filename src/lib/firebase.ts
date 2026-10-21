import { initializeApp, getApps, getApp, type FirebaseApp } from "firebase/app";
import {
  getFirestore,
  type Firestore,
  collection,
  doc,
  getDocs,
  getDoc,
  setDoc,
  deleteDoc,
  onSnapshot,
  query,
  orderBy,
  type Unsubscribe,
} from "firebase/firestore";
import firebaseConfig from "../../firebase-applet-config.json";

// Configuração oficial do Firebase provisionada para o projeto
const config = {
  apiKey: firebaseConfig.apiKey,
  authDomain: firebaseConfig.authDomain,
  projectId: firebaseConfig.projectId,
  storageBucket: firebaseConfig.storageBucket,
  messagingSenderId: firebaseConfig.messagingSenderId,
  appId: firebaseConfig.appId,
};

const DATABASE_ID =
  firebaseConfig.firestoreDatabaseId && firebaseConfig.firestoreDatabaseId !== "(default)"
    ? firebaseConfig.firestoreDatabaseId
    : "(default)";

let appInstance: FirebaseApp | null = null;
let dbInstance: Firestore | null = null;

export function getFirebaseApp(): FirebaseApp {
  if (!appInstance) {
    if (getApps().length > 0) {
      appInstance = getApp();
    } else {
      appInstance = initializeApp(config);
    }
  }
  return appInstance;
}

export function getFirebaseFirestore(): Firestore {
  if (!dbInstance) {
    const app = getFirebaseApp();
    dbInstance = getFirestore(app, DATABASE_ID);
  }
  return dbInstance;
}

export const db = getFirebaseFirestore();

export {
  collection,
  doc,
  getDocs,
  getDoc,
  setDoc,
  deleteDoc,
  onSnapshot,
  query,
  orderBy,
  type Unsubscribe,
};
