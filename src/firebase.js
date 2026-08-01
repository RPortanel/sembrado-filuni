import { initializeApp } from "firebase/app";
import { getFirestore } from "firebase/firestore";

// REEMPLAZA ESTO CON TUS CREDENCIALES EXACTAS
const firebaseConfig = {
  apiKey: "AIzaSyDBd8ZDF2fFeuxvaliNO_OweiR_6gAo7SM",
  authDomain: "filuni-sembrado-2026.firebaseapp.com",
  projectId: "filuni-sembrado-2026",
  storageBucket: "filuni-sembrado-2026.firebasestorage.app",
  messagingSenderId: "381662574849",
  appId: "1:381662574849:web:99052805602ed120768300"
};

// Inicializar Firebase y la Base de Datos
const app = initializeApp(firebaseConfig);
export const db = getFirestore(app);
