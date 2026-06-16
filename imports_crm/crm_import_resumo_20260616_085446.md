# Resumo da conversao

- Data de geracao: 2026-06-16 08:54:47
- Planilhas processadas: Controle de CRM 2026 (1).xlsx e Controle de Contratos Gerados pelo Marketing 2026 (1).xlsx
- Linhas convertidas: 787
- Linhas recomendadas para importacao apos deduplicacao: 779
- Duplicidades removidas no arquivo recomendado: 8
- Linhas puladas por falta de nome ou contato valido: 24

## Arquivos

- Todos os registros convertidos: crm_import_todos_20260616_085446.csv
- Arquivo recomendado para importar: crm_import_deduplicado_20260616_085446.csv
- Auditoria completa: crm_import_auditoria_20260616_085446.csv
- Duplicidades encontradas: crm_import_duplicados_20260616_085446.csv
- Linhas puladas: crm_import_pulados_20260616_085446.csv

## Regras aplicadas

- Telefones foram normalizados para apenas numeros e DDD 64 foi aplicado quando o numero tinha 8 ou 9 digitos.
- Datas foram convertidas para o formato AAAA-MM-DD.
- origem foi padronizada em Organico ou Pago para manter compatibilidade com o CRM atual.
- pipeline foi inferido a partir do status original com os valores Novos Leads, Em Contato, Proposta, Fechado e Cancelado.
- observacoes mantem a rastreabilidade da linha original e os detalhes da planilha.