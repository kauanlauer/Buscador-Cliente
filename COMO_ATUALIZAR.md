**Fluxo**

1. Gere a nova versão:
```powershell
powershell -NoProfile -ExecutionPolicy Bypass -File .\build_instalador_buscador.ps1
```

2. Publique a atualização em uma pasta compartilhada, OneDrive ou rede:
```powershell
powershell -NoProfile -ExecutionPolicy Bypass -File .\publicar_atualizacao_buscador.ps1 -Version 1.0.2 -Destino "D:\PastaCompartilhada\Buscador"
```

3. Isso vai criar dois arquivos na pasta de destino:
- `Buscador Cliente HeadCargo.exe`
- `manifesto_buscador_cliente_headcargo.json`

4. Em cada máquina do usuário, configure uma vez o caminho do manifesto em:
- `Configuracoes`
- `Manifesto de atualizacao (.json)`

5. Depois disso, o usuário atualiza pelo próprio programa:
- botão `Atualizar`
- ou menu `Buscador > Verificar atualizacao`

**Observacoes**

- O manifesto pode ficar em uma pasta do OneDrive da empresa.
- Sempre que você publicar uma versão nova com número maior, os usuários conseguem atualizar sem reinstalar.
- O instalador completo continua disponível quando você quiser instalar do zero em outra máquina.
