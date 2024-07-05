SELECT
CODIGO Numero_Prontuario,
NOME_COMPLETO Nome_Completo, --Se estiver em branco, então pegue NOME,
NOME Apelido, --.GetPrimeirosCaracteres(20).ToNome(); -- Se estiver em branco, deixa em branco,
OBS Observacoes,
EMAIL E_mail,
CRO Codigo_do_Conselho_e_Estado,
TELEFONE Telefone_Principal
FROM CED006 