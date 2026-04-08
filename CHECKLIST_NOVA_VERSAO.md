**Checklist De Nova Versao**

Este arquivo existe para qualquer proxima IA ou pessoa saber exatamente o que precisa ser atualizado quando sair uma nova versao do Buscador Cliente HeadCargo.

**Quando for lancar uma nova versao**

1. Atualize a versao do app em:
- [launcher_clientes_onedrive.pyw](/C:/Users/Kauan%20Lauer/Documents/Scripts/Buscador/launcher_clientes_onedrive.pyw)
  Procure por `APP_VERSION = "X.Y.Z"`

2. Atualize a versao do instalador em:
- [buscador_cliente_headcargo.iss](/C:/Users/Kauan%20Lauer/Documents/Scripts/Buscador/buscador_cliente_headcargo.iss)
  Procure por `#define MyAppVersion "X.Y.Z"`

3. Gere os executaveis:
```powershell
powershell -NoProfile -ExecutionPolicy Bypass -File .\build_instalador_buscador.ps1
```

4. Atualize o manifesto do GitHub com a nova versao e uma nota curta:
```powershell
powershell -NoProfile -ExecutionPolicy Bypass -File .\publicar_atualizacao_buscador.ps1 -Version X.Y.Z -Notes "Descricao curta da versao."
```

5. Monte a pasta `Github` com os arquivos que vao para o repositorio:
```powershell
powershell -NoProfile -ExecutionPolicy Bypass -File .\preparar_pasta_github.ps1
```

6. Suba para o repositorio GitHub somente os arquivos que estiverem dentro da pasta `Github`.

7. Crie um novo `Release` no GitHub e anexe este arquivo:
- `dist\Setup Buscador Cliente HeadCargo.exe`

O nome do arquivo no Release deve continuar exatamente:
- `Setup Buscador Cliente HeadCargo.exe`

**Importante**

- O app baixa atualizacao deste manifesto:
  `https://raw.githubusercontent.com/kauanlauer/Buscador-Cliente/main/github_update_manifest.json`
- O app baixa o instalador deste link:
  `https://github.com/kauanlauer/Buscador-Cliente/releases/latest/download/Setup%20Buscador%20Cliente%20HeadCargo.exe`
- Se nao existir Release com esse arquivo, a atualizacao automatica nao funciona.

**Nao subir no repositorio**

- `dist`
- `build`
- `__pycache__`
- `launcher_clientes_onedrive_config.json`
- arquivos `Zone.Identifier`

**Observacao**

- Se voce mexeu no codigo mas ainda nao vai publicar para usuarios, nao precisa trocar a versao.
