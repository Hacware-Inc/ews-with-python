
$hello_variable = 'Hello Word'
Write-Output $hello_variable

if($hello_variable -like "H*"){
  Write-Output 'H Found'
}

for ($i -in $hello_variable){
  Write-Output $hello_variable
}
