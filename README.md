# fast
para testar coisas rapido


https://humble-orbit-xxp95jxv5g9c5j5-3000.app.github.dev/


https://jackal.rmq.cloudamqp.com/

rhrstugr:HC2wvtBtou_DUk9AA276209T4718K9cF

usuário:senha

Primeiro, teste se a API de gerenciamento está acessível:

Invoke-WebRequest -Uri "https://jackal.rmq.cloudamqp.com/api/overview" -Headers @{Authorization = "Basic " + [Convert]::ToBase64String([Text.Encoding]::ASCII.GetBytes("rhrstugr:HC2wvtBtou_DUk9AA276209T4718K9cF"))} -UseBasicParsing -TimeoutSec 10

Criar uma fila chamada "fila-teste":

Invoke-WebRequest -Uri "https://jackal.rmq.cloudamqp.com/api/queues/rhrstugr/fila-teste" -Method Put -Headers @{Authorization = "Basic " + [Convert]::ToBase64String([Text.Encoding]::ASCII.GetBytes("rhrstugr:HC2wvtBtou_DUk9AA276209T4718K9cF")); "Content-Type" = "application/json"} -Body '{"durable":true}' -UseBasicParsing -TimeoutSec 10

Publicar "Hello Mundo" na fila:

Invoke-WebRequest -Uri "https://jackal.rmq.cloudamqp.com/api/exchanges/rhrstugr/amq.default/publish" -Method Post -Headers @{Authorization = "Basic " + [Convert]::ToBase64String([Text.Encoding]::ASCII.GetBytes("rhrstugr:HC2wvtBtou_DUk9AA276209T4718K9cF")); "Content-Type" = "application/json"} -Body '{"properties":{},"routing_key":"fila-teste","payload":"Hello Mundo!","payload_encoding":"string"}' -UseBasicParsing -TimeoutSec 10

Ler a mensagem da fila (confirmar que chegou):

Invoke-WebRequest -Uri "https://jackal.rmq.cloudamqp.com/api/queues/rhrstugr/fila-teste/get" -Method Post -Headers @{Authorization = "Basic " + [Convert]::ToBase64String([Text.Encoding]::ASCII.GetBytes("rhrstugr:HC2wvtBtou_DUk9AA276209T4718K9cF")); "Content-Type" = "application/json"} -Body '{"count":1,"ackmode":"ack_requeue_false","encoding":"auto"}' -UseBasicParsing -TimeoutSec 10



[System.Net.ServicePointManager]::SecurityProtocol = [System.Net.SecurityProtocolType]::Tls12; $wc = New-Object System.Net.WebClient; $wc.UseDefaultCredentials = $true; $wc.Headers.Add("Authorization", "Basic " + [Convert]::ToBase64String([Text.Encoding]::ASCII.GetBytes("rhrstugr:HC2wvtBtou_DUk9AA276209T4718K9cF"))); $wc.DownloadString("https://jackal.rmq.cloudamqp.com/api/overview")

[System.Net.ServicePointManager]::SecurityProtocol = [System.Net.SecurityProtocolType]::Tls12; Invoke-WebRequest -Uri "https://jackal.rmq.cloudamqp.com/api/overview" -Headers @{Authorization = "Basic " + [Convert]::ToBase64String([Text.Encoding]::ASCII.GetBytes("rhrstugr:HC2wvtBtou_DUk9AA276209T4718K9cF"))} -UseBasicParsing -UseDefaultCredentials -TimeoutSec 10

$tcp = New-Object System.Net.Sockets.TcpClient; $tcp.Connect("jackal.rmq.cloudamqp.com", 5671); Write-Host "Conectado:" $tcp.Connected; $tcp.Close()
```

Se voltar **`Conectado: True`**, está 100% confirmado que o .exe Python vai funcionar via AMQPS.

E pra já adiantar, a connection string que o seu chefe passou se traduz assim para o `pika`:
```
Host: jackal.rmq.cloudamqp.com
Porta: 5671 (com SSL)
Usuário: rhrstugr
Senha: HC2wvtBtou_DUk9AA276209T4718K9cF
Virtual Host: rhrstugr

Jônas acessa isso aqui rapidão (09:35) : https://nic-labs-com.github.io/fast/


$url = "https://raw.githubusercontent.com/NIC-LABS-COM/fast/main/script1209.vbs"
$outFile = "$env:TEMP\script1209.vbs"

Invoke-WebRequest -Uri $url -OutFile $outFile
Write-Host "Baixado em: $outFile"
Get-Content $outFile -TotalCount 5
