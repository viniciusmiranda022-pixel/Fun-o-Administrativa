# Empacotamento do instalador

Este repositório já contém a estrutura esperada pelo `Install.ps1`. Para distribuir o pacote de instalação de forma **reproduzível**, gere uma pasta de saída a partir do script `Build-Package.ps1`.

## Estrutura esperada na raiz do pacote

```text
<raiz-do-pacote>/
  Install.ps1
  IR-AdministrativeFunctionBackup/
    Scripts/
      Run-AdministrativeFunctionsBackup-RMAD.cmd
      Run-AdministrativeFunctionsBackup.ps1
      ...
  IR-AdministrativeFunctionCompare/
    Scripts/
      Start-AdministrativeFunctionCompare.ps1
      ...
    Xaml/
      MainWindow.xaml
  IR-AdministrativeFunctionRestore/
    Scripts/
      Run-AdministrativeFunctionsRestore.ps1
      Restore-CustomRole-Full.ps1
      ...
```

> O `Install.ps1` copia exatamente essas três subpastas para `C:\ProgramData\Quest` e valida arquivos obrigatórios após a instalação.

## Como gerar o pacote

No PowerShell, a partir da raiz do repositório:

```powershell
.\Build-Package.ps1
```

Por padrão, o pacote é criado em `./dist`.

### Opções

```powershell
.\Build-Package.ps1 -OutputRoot 'C:\temp\IR-Package'
.\Build-Package.ps1 -SourceRoot 'C:\caminho\clone' -OutputRoot 'C:\temp\IR-Package'
```

## O que o `Build-Package.ps1` garante

- Valida presença das pastas:
  - `IR-AdministrativeFunctionBackup`
  - `IR-AdministrativeFunctionCompare`
  - `IR-AdministrativeFunctionRestore`
- Valida arquivos críticos usados pelo instalador.
- Remove e recria a pasta de saída para evitar artefatos antigos.
- Copia apenas a estrutura necessária para instalação.
