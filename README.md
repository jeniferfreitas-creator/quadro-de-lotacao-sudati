# Quadro de Lotação — Sudati

Site do Quadro de Lotação, Organograma, Vagas em Aberto e Indicadores.

Os dados ficam salvos num banco de dados na nuvem (Firebase/Firestore,
gratuito), para que todas as pessoas que acessarem o site vejam as
mesmas informações.

Veja o passo a passo completo de publicação na mensagem que a Claude te
enviou, ou siga o resumo abaixo.

## Resumo rápido

1. Crie um projeto gratuito em https://console.firebase.google.com,
   ative o Firestore Database, e copie as chaves para `src/firebaseConfig.js`
2. Crie uma conta gratuita em https://github.com
3. Crie um repositório novo chamado `quadro-de-lotacao-sudati`
4. Suba os arquivos desta pasta para esse repositório
5. Em Settings → Pages, escolha "GitHub Actions" como fonte
6. Aguarde alguns minutos — o site fica em:
   `https://SEU-USUARIO.github.io/quadro-de-lotacao-sudati/`

Todo push feito na branch `main` publica uma nova versão automaticamente.
