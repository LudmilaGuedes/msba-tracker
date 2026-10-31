"""
Politica de retencao para msba_programs_database.xlsx.

Regra: no maximo 2 versoes do banco de dados ficam em data/ a
qualquer momento --
  1. msba_programs_database.xlsx        (versao atual/vigente)
  2. msba_programs_database_backup.xlsx (a versao imediatamente anterior)

Nomenclatura fixa (Opcao A), nao timestamped: o dashboard React le sempre
o mesmo caminho estatico (public/data/msba_programs_database.xlsx), entao
manter o nome do arquivo principal fixo evita ter que tornar a leitura do
frontend dinamica (o que exigiria um backend/API para listar arquivos --
algo que o app estatico atual nao tem).

Uso em qualquer script que vai sobrescrever o banco de dados:

    from db_retention import rotate_backup, DB_PATH

    rotate_backup()          # chamar ANTES de abrir/gravar o novo conteudo
    wb = openpyxl.load_workbook(DB_PATH)
    ...
    wb.save(DB_PATH)
"""

import shutil
from pathlib import Path

DB_PATH = Path("data/msba_programs_database.xlsx")
BACKUP_PATH = Path("data/msba_programs_database_backup.xlsx")


def rotate_backup() -> None:
    """Move a versao atual para o slot unico de backup, descartando
    qualquer backup anterior. Chamar sempre ANTES de escrever novos dados
    em DB_PATH. Garante que o diretorio nunca tenha mais que essas 2
    versoes do banco de dados."""
    if not DB_PATH.exists():
        return
    if BACKUP_PATH.exists():
        BACKUP_PATH.unlink()
    shutil.copy2(DB_PATH, BACKUP_PATH)


def enforce_retention() -> list[str]:
    """Escaneia data/ por variantes antigas do banco de dados
    (o padrao timestamped 'msba_programs_database.backup-*.xlsx' usado
    antes desta politica) e apaga tudo exceto a mais recente, que e
    renomeada para o nome fixo de backup. Retorna a lista de arquivos
    apagados. Serve para limpar o diretorio uma vez, apos adotar esta
    politica; scripts novos devem usar rotate_backup() em vez disto."""
    cwd = Path("data")
    legacy = sorted(
        cwd.glob("msba_programs_database.backup-*.xlsx"),
        key=lambda p: p.stat().st_mtime,
    )
    deleted = []
    if not legacy:
        return deleted

    most_recent = legacy[-1]
    for old in legacy[:-1]:
        old.unlink()
        deleted.append(str(old))

    if BACKUP_PATH.exists():
        BACKUP_PATH.unlink()
        deleted.append(str(BACKUP_PATH) + " (substituído)")
    most_recent.rename(BACKUP_PATH)

    return deleted


if __name__ == "__main__":
    removed = enforce_retention()
    if removed:
        print(f"Removidos/consolidados {len(removed)} arquivo(s) legado(s):")
        for f in removed:
            print(f"  - {f}")
        print(f"Backup consolidado em: {BACKUP_PATH}")
    else:
        print("Nenhum backup legado encontrado. Nada a fazer.")
