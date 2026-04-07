**Fluxo GitHub**

1. Gere a nova versao:
```powershell
powershell -NoProfile -ExecutionPolicy Bypass -File .\build_instalador_buscador.ps1
```

2. Atualize o manifesto do repositorio:
```powershell
powershell -NoProfile -ExecutionPolicy Bypass -File .\publicar_atualizacao_buscador.ps1 -Version 1.1.0 -Notes "Descricao curta da versao."
```

3. Envie para o GitHub:
- `github_update_manifest.json`
- commit e push no repositorio

4. Crie ou atualize o GitHub Release e anexe exatamente este arquivo:
- `dist\Setup Buscador Cliente HeadCargo.exe`

O nome do anexo no Release deve ser:
- `Setup Buscador Cliente HeadCargo.exe`

**Como o app atualiza**

- O programa verifica atualizacao automaticamente 1 vez por dia.
- O usuario tambem pode clicar em `Verificar atualizacao`.
- Quando existe versao nova, o app pergunta se deseja atualizar.
- Se confirmar, ele baixa o `Setup` direto do GitHub e instala em silencio.
- A configuracao atual do usuario e da pasta dos clientes e preservada no update.

**Observacoes**

- O app busca o manifesto em:
  `https://raw.githubusercontent.com/kauanlauer/Buscador-Cliente/main/github_update_manifest.json`
- O instalador e baixado de:
  `https://github.com/kauanlauer/Buscador-Cliente/releases/latest/download/Setup%20Buscador%20Cliente%20HeadCargo.exe`
- Se o Release estiver sem esse arquivo, o update nao vai conseguir baixar.
