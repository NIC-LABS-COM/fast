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
