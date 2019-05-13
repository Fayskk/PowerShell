#Variavél para inserir nome do usuário:
$NomeUsuario = "alessandro.nogueira","alessandro.nogueira"
#Variavél para incerir nome do computador:
$NomePc = "oficinadeti","oficinadeti"
#Comando FOR para rodar os scipts enquanto existir usuário:
for ($i = 0; $i -lt $NomeUsuario.Count; $i += 1) {
    $PC = $NomePc[$i] + "\c$\Users\" + $NomeUsuario[$i] + "\Documents\Arquivos do Outlook"
    $pc
    # Criando caminho mapeado para W:
    New-PSDrive -Name "W" -Root "\\$PC" -Persist -PSProvider "FileSystem"
    # Copiando arquivos PST: 
    Copy-Item -Path "W:\*" -Destination (new-item -type directory -force ("C:\Drawings\" + $NomeUsuario[$i])) -Force -ea 0
    # Exluindo caminho mapeado para W:
    Get-PSDrive W | Remove-PSDrive
 }
#Fim


 
 #copy-item -Path $file -Destination (new-item -type directory -force ("C:\Folder\sub\sub\" + $newSub)) -force 