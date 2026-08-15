import { defineConfig } from "vite";
import react from "@vitejs/plugin-react";

// IMPORTANT: "base" precisa ser "/NOME-DO-SEU-REPOSITORIO/"
// Se você seguiu o passo a passo e criou o repositório com o nome
// "quadro-de-lotacao-sudati", pode deixar como está.
// Se usou outro nome, troque abaixo pelo nome exato do seu repositório.
export default defineConfig({
  plugins: [react()],
  base: "/quadro-de-lotacao-sudati/",
});
