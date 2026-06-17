from __future__ import annotations

import csv
import re
from collections import Counter
from pathlib import Path

BASE = Path(r"c:\GitHub\CRMPaxRioVerde\CRMPaxRioVerde\imports_crm")
INPUT = BASE / "crm_import_deduplicado_20260616_145558.csv"
OUTPUT = BASE / "crm_import_deduplicado_20260616_151000.csv"
SUMMARY = BASE / "crm_import_resumo_20260616_151000.txt"

PIPE_FOLLOW_UP = "Follow-Up"
PIPE_NEGOTIATION = "Apresentação/Negociação"
PIPE_NO_RESPONSE = "Não responde"
PIPE_CLOSED = "Fechado"
PIPE_ALREADY_ASSOCIATED = "Já é associado"
PIPE_NO_INTEREST = "Não tem interesse"
PIPE_CANNOT_REACTIVATE = "Cliente não pode reativar o plano"

SPLIT_RE = re.compile(r"\s*\|\|\s*")


def normalize_spaces(text: str) -> str:
    return re.sub(r"\s+", " ", (text or "").replace("\xa0", " ")).strip()


def normalize_lookup(text: str) -> str:
    import unicodedata

    value = normalize_spaces(text)
    value = unicodedata.normalize("NFD", value)
    value = "".join(ch for ch in value if unicodedata.category(ch) != "Mn")
    return value.lower()


def unique_keep_order(items: list[str]) -> list[str]:
    seen: set[str] = set()
    result: list[str] = []
    for item in items:
        value = normalize_spaces(item)
        key = normalize_lookup(value)
        if not value or key in seen:
            continue
        seen.add(key)
        result.append(value)
    return result


def extract_meaningful_fragments(note: str) -> list[str]:
    note = normalize_spaces(note)
    if not note:
        return []

    fragments: list[str] = []
    label_patterns = [
        r"Detalhes:\s*([^|]+)",
        r"Status 1:\s*([^|]+)",
        r"Status original:\s*([^|]+)",
        r"Status do atendimento:\s*([^|]+)",
        r"Status final:\s*([^|]+)",
        r"Motivo nao fechamento:\s*([^|]+)",
    ]

    for pattern in label_patterns:
        for match in re.finditer(pattern, note, flags=re.IGNORECASE):
            fragments.append(match.group(1).strip())

    if fragments:
        return unique_keep_order(fragments)

    pieces = [p.strip() for p in note.split("|")]
    ignored_prefixes = (
        "origem:",
        "canal original:",
        "cidade:",
        "contato pessoa:",
        "mes referencia:",
        "dia indicacao:",
        "lider:",
        "quantidade original:",
        "origem detalhada:",
        "data fechamento:",
        "ja e cliente:",
    )
    kept = [
        piece
        for piece in pieces
        if piece and not normalize_lookup(piece).startswith(ignored_prefixes)
    ]
    return unique_keep_order(kept)


def clean_observations(raw: str) -> str:
    groups = SPLIT_RE.split(raw or "")
    cleaned: list[str] = []
    for group in groups:
        cleaned.extend(extract_meaningful_fragments(group))
    return " | ".join(unique_keep_order(cleaned))


def infer_pipeline(text: str) -> str:
    lookup = normalize_lookup(text)
    if not lookup:
      return PIPE_FOLLOW_UP

    closed = [
        r"\bfechado\b",
        r"\bfechou\b",
        r"\bfeito\b",
        r"fez o contrato",
        r"fez pela empresa",
        r"realizou o plano",
        r"ja fez o plano",
        r"contratos empresariais",
    ]
    if any(re.search(p, lookup) for p in closed):
        return PIPE_CLOSED

    cannot_reactivate = [
        r"nao pode reativa",
        r"nao pode reativar",
        r"nao consegue reativar",
        r"nao seria possivel reativar",
    ]
    if any(re.search(p, lookup) for p in cannot_reactivate):
        return PIPE_CANNOT_REACTIVATE

    already_associated = [
        r"ja e cliente",
        r"ja e associado",
        r"cliente da pax social",
        r"ja esta associado",
        r"ja tem plano",
        r"ja possui plano",
    ]
    if any(re.search(p, lookup) for p in already_associated):
        return PIPE_ALREADY_ASSOCIATED

    no_interest = [
        r"nao tem interesse",
        r"nao tem interresse",
        r"sem interesse",
        r"sem interresse",
        r"cliente sem interesse",
        r"nao quer",
        r"perdido",
        r"encerrado",
        r"ja foi cliente",
        r"mudou para outra cidade",
        r"mudou para santa helena",
        r"nao mora",
        r"diretoria barrou",
        r"barrou a proposta",
        r"nao sendo possivel fazer o plano",
        r"nao e possivel fazer o plano",
    ]
    if any(re.search(p, lookup) for p in no_interest):
        return PIPE_NO_INTEREST

    no_response = [
        r"nao responde",
        r"nao reponde",
        r"nao atende",
        r"sem resposta",
        r"nao visualiza",
        r"numero errado",
        r"n errado",
        r"telefone inexistente",
        r"numero de telefone ixesistente",
        r"numero nao chama",
        r"nao possui whatsapp",
        r"nao tem whatsapp",
        r"nao tem zap",
        r"nao possui zap",
        r"ultimo contato",
        r"ultima tentativa",
        r"varias tentativas",
        r"caixa postal",
        r"mensagens temporarias",
    ]
    if any(re.search(p, lookup) for p in no_response):
        return PIPE_NO_RESPONSE

    negotiation = [
        r"proposta",
        r"apresentad",
        r"apresentar",
        r"reuniao",
        r"marcar um dia",
        r"marcar um horario",
        r"marcar horario",
        r"agend",
        r"vir ao escritorio",
        r"quer fazer",
        r"quer fechar",
        r"vai fechar",
        r"vai analisar",
        r"indeciso",
        r"interessado",
        r"interesse",
        r"interresse",
        r"ficou de acertar",
        r"ficou de ver",
        r"ficou de falar",
        r"avaliando a possibilidade",
        r"priorizar outros assuntos",
    ]
    if any(re.search(p, lookup) for p in negotiation):
        return PIPE_NEGOTIATION

    follow_up = [
        r"em andamento",
        r"cliente esta com o vendedor",
        r"liguei",
        r"\bligar\b",
        r"\bmsg\b",
        r"mensagem",
        r"retorno",
        r"retornar",
        r"aguardando resposta",
        r"aguardando retorno",
        r"aguardando",
        r"vai falar com",
        r"enviar e-mail",
        r"enviado e-mail",
        r"material explicativo",
        r"vamos entrar em contato",
        r"vamos ligar",
        r"ligar na semana",
        r"ligar segunda",
        r"passou o contato",
        r"entrar em contato",
    ]
    if any(re.search(p, lookup) for p in follow_up):
        return PIPE_FOLLOW_UP

    return PIPE_FOLLOW_UP


with INPUT.open("r", encoding="utf-8-sig", newline="") as fh:
    rows = list(csv.DictReader(fh, delimiter=";"))
    fieldnames = list(rows[0].keys()) if rows else []

for row in rows:
    row["origem"] = row["origem"].replace("IndicaÃ§Ã£o", "Indicação")
    if normalize_lookup(row["origem"]) == normalize_lookup("IndicaÃ§Ã£o"):
        row["origem"] = "Indicação"
    if normalize_lookup(row["origem"]) == normalize_lookup("Indicacao"):
        row["origem"] = "Indicação"
    row["observacoes"] = clean_observations(row.get("observacoes", ""))
    row["pipeline"] = infer_pipeline(" | ".join([row.get("pipeline", ""), row.get("observacoes", "")]))

# repair any remaining mojibake source values
for row in rows:
    if normalize_lookup(row["origem"]) in {normalize_lookup("Indicacao"), normalize_lookup("IndicaÃ§Ã£o")}:
        row["origem"] = "Indicação"
    if normalize_lookup(row["rede_social"]) == normalize_lookup("IndicaÃ§Ã£o Colaborador"):
        row["rede_social"] = "Indicação Colaborador"

with OUTPUT.open("w", encoding="utf-8-sig", newline="") as fh:
    writer = csv.DictWriter(fh, fieldnames=fieldnames, delimiter=";")
    writer.writeheader()
    writer.writerows(rows)

pipeline_counts = Counter(row["pipeline"] for row in rows)
summary_lines = [
    f"Arquivo: {OUTPUT.name}",
    f"Linhas: {len(rows)}",
    *[f"{name}: {count}" for name, count in sorted(pipeline_counts.items())],
]
SUMMARY.write_text("\n".join(summary_lines), encoding="utf-8")
