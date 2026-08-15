import React from "react";
import ReactDOM from "react-dom/client";
import App from "./App.jsx";
import { doc, getDoc, setDoc, deleteDoc, collection, getDocs } from "firebase/firestore";
import { db } from "./firebase.js";

// ---------------------------------------------------------------------
// Polyfill do "window.storage" (API usada quando este painel rodava como
// artefato dentro do Claude), agora ligado ao Firestore — um banco de
// dados na nuvem gratuito do Google. Assim, TODAS as pessoas que
// acessarem o site veem os mesmos dados, porque tudo fica salvo num
// lugar só, não mais no navegador de cada pessoa.
//
// Não precisa mexer neste arquivo — só preencher o firebaseConfig.js
// com as chaves do seu projeto Firebase (veja o passo a passo).
// ---------------------------------------------------------------------
const COLLECTION = "quadro-lotacao-dados";

window.storage = {
  async get(key) {
    const snap = await getDoc(doc(db, COLLECTION, key));
    if (!snap.exists()) {
      throw new Error(`Chave "${key}" não encontrada`);
    }
    return { key, value: snap.data().value, shared: true };
  },
  async set(key, value) {
    await setDoc(doc(db, COLLECTION, key), { value });
    return { key, value, shared: true };
  },
  async delete(key) {
    await deleteDoc(doc(db, COLLECTION, key));
    return { key, deleted: true, shared: true };
  },
  async list(prefix) {
    const snaps = await getDocs(collection(db, COLLECTION));
    const keys = snaps.docs.map((d) => d.id).filter((k) => !prefix || k.startsWith(prefix));
    return { keys, prefix, shared: true };
  },
};

ReactDOM.createRoot(document.getElementById("root")).render(
  <React.StrictMode>
    <App />
  </React.StrictMode>
);
