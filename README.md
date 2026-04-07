**GitHub**

Use esta pasta como referencia rapida do que vai para o GitHub.

**Repositorio**

Arquivos que devem estar no repositorio:
- `launcher_clientes_onedrive.pyw`
- `build_instalador_buscador.ps1`
- `gerar_icone_buscador.ps1`
- `buscador_cliente_headcargo.iss`
- `publicar_atualizacao_buscador.ps1`
- `github_update_manifest.json`
- `COMO_ATUALIZAR.md`
- `logo_buscador.png`
- `logo_buscador.ico`
- `Buscador Cliente HeadCargo.spec`
- `.gitignore`

**Release**

Arquivo que deve ir no GitHub Release:
- `dist\Setup Buscador Cliente HeadCargo.exe`

**Como organizar**

1. Gere a nova versao com o build.
2. Atualize o manifesto com `publicar_atualizacao_buscador.ps1`.
3. Rode `Github\preparar_publicacao.ps1`.
4. Use a pasta `Github\Repositorio` como checklist do que vai para commit/push.
5. Use a pasta `Github\Release` como checklist do que vai para o Release.
