###################################################
# @file to.ps1
# @brief 将图片集转化为 Manager.py 可以识别的图片库
# @note 只处理 PNG 文件
# @warning 只能执行一次，执行前请做好备份
# @author Rikka Github/ming-14
###################################################

$jsonArray = @()

Get-ChildItem -Path . -Filter *.png | ForEach-Object {
    $hash = (Get-FileHash -Path $_.FullName -Algorithm SHA256).Hash
    $brief = $_.BaseName
    $newName = "$hash.pngx" # 构建新文件名
    
    $jsonArray += [PSCustomObject]@{
        name   = $newName
        sha256 = $hash
        brief  = $brief
    }
    
    # 重命名文件
    Rename-Item -Path $_.FullName -NewName $newName
}

# 写入 JSON
$jsonArray | ConvertTo-Json | Out-File "files.json" -Encoding utf8
