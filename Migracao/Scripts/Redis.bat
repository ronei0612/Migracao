:: LIST Funcionarios
redis-cli.exe -h 127.0.0.1 -n 7 get Equipe:099999-Funcionarios
:: CLEAR Funcionarios
redis-cli.exe -h 127.0.0.1 -n 7 del Equipe:099999-Funcionarios

:: LIST PrecosTabelas
redis-cli.exe -h 127.0.0.1 -n 0 get Tabelas:Tabelas-099999
:: CLEAR PrecosTabelas
redis-cli.exe -h 127.0.0.1 -n 0 del Tabelas:Tabelas-099999