import { initializeApp } from "firebase/app";
import { getFirestore } from "firebase/firestore";

const firebaseConfig = {
  apiKey: "AIzaSyAa9TJs1jMUACHzNcCbl6rCXy--WStCQSo",
  authDomain: "inedisav-formatos.firebaseapp.com",
  projectId: "inedisav-formatos",
  storageBucket: "inedisav-formatos.firebasestorage.app",
  messagingSenderId: "665501882360",
  appId: "1:665501882360:web:6c5e6079ca38af23bce80d"
};

const app = initializeApp(firebaseConfig);
export const db = getFirestore(app);
