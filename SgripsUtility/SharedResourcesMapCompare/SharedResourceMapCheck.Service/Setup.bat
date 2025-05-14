echo off
sc stop EasySpedWSInterfaceService
sc delete EasySpedWSInterfaceService
echo Installazione Servizio DB EasySpedDE
echo SetupEasySpedWSInterfaceServiceUpdate.msi /qb 