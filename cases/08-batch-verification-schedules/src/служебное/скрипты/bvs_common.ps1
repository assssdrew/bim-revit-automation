# UTF-8 WITH BOM required (PowerShell 5.1).
# Общие функции, которые раньше копировались в несколько скриптов сборки.
# Подключать через точку. Если файл не подключился, у вызывающего остаётся запасное определение.

function ConvertTo-CanonicalSpecName([string]$Name) {
    return ([string]$Name).Replace("значния", "значения").Replace(([string][char]0x43 + "антехприборы"), ([string][char]0x421 + "антехприборы")).Replace(([string][char]0x63 + "антехприборы"), ([string][char]0x441 + "антехприборы"))
}
