# Resumo da conversao

- Data de geracao: 2026-06-16 15:21:16
- Planilhas processadas: Controle de CRM 2026 (1).xlsx, Controle de Contratos Gerados pelo Marketing 2026 (1).xlsx e PROGRAMA DE INDICAÇÕES COLABORADOR.xlsx
- Linhas convertidas: 1276
- Linhas recomendadas para importacao apos deduplicacao: 1255
- Duplicidades removidas no arquivo recomendado: 21
- Linhas puladas por falta de nome ou contato valido: 35

## Arquivos

- Todos os registros convertidos: crm_import_todos_20260616_152115.csv
- Arquivo recomendado para importar: crm_import_deduplicado_20260616_152115.csv
- Auditoria completa: crm_import_auditoria_20260616_152115.csv
- Duplicidades encontradas: crm_import_duplicados_20260616_152115.csv
- Linhas puladas: crm_import_pulados_20260616_152115.csv

## Regras aplicadas

- Telefones foram normalizados para apenas numeros e DDD 64 foi aplicado quando o numero tinha 8 ou 9 digitos.
- Datas foram convertidas para o formato AAAA-MM-DD.
- origem foi padronizada em IndicaÃ§Ã£o, Evento Externo, Organico ou Pago para manter compatibilidade com o CRM atual.
- pipeline foi inferido a partir do status original com os valores Novos Leads, Em Contato, Proposta, Fechado e Cancelado.
- observacoes mantem a rastreabilidade da linha original e os detalhes da planilha.