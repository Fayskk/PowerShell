﻿#Variavél para inserir nome do usuário:
$NomeUsuario = "alessandro.nogueira","william.filso"
#Variavél para incerir nome do computador:
$NomePc = "oficinadeti","WILLIAN-DESK"
#Comando FOR para rodar os scipts enquanto existir usuário:
for ($i = 0; $i -lt $NomeUsuario.Count; $i += 1) {
    $PC = $NomePc[$i] + "\c$\Correio"
    $pc
    # Criando caminho mapeado para W:
    New-PSDrive -Name "W" -Root "\\$PC" -Persist -PSProvider "FileSystem"
    # Copiando arquivos PST: 
    Copy-Item -Path "W:\*" -Destination (new-item -type directory -force ("C:\Drawings\" + $NomeUsuario[$i])) -Force -ea 0
    # Exluindo caminho mapeado para W:
    Get-PSDrive W | Remove-PSDrive
 }
#Fim