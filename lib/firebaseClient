import { initializeApp, getApps } from "firebase/app";
import { getAuth } from "firebase/auth";
import { getFirestore } from "firebase/firestore";

const firebaseConfig = {
  apiKey: "AIzaSyAvZqK2BNpHJ8hqKzfz_yXAzaKNnu0_lqE",
  authDomain: "exp-cs-acocollo-i-1.firebaseapp.com",
  projectId: "exp-cs-acocollo-i-1",
  storageBucket: "exp-cs-acocollo-i-1.firebasestorage.app",
  messagingSenderId: "914924991021",
  appId: "1:914924991021:web:5bc72b11f109a39cd3bea4",
  measurementId: "G-HBN3J4D98D"
};

const app = !getApps().length ? initializeApp(firebaseConfig) : getApps()[0];
export const auth = getAuth(app);
export const db = getFirestore(app);
