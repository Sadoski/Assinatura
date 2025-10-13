# Caminho da imagem
$imagePath = ".\\imagem.png"

# Ler a imagem como um array de bytes
$imageBytes = [System.IO.File]::ReadAllBytes($imagePath)

# Converter o array de bytes para Base64
$base64String = [Convert]::ToBase64String($imageBytes)

# Criar a string Base64 com o prefixo 'data:image/png;base64,' para uso em HTML
$base64Data = "data:image/png;base64,$base64String"

# Salvar o resultado em um arquivo de texto
Set-Content -Path ".\dados.txt" -Value $base64Data -Encoding UTF8

# Exibir os primeiros 200 caracteres para visualização
$base64Data.Substring(0, 200)
