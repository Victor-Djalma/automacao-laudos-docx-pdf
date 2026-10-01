# Automação de Laudos DOCX → PDF

Automação em Python para organizar lotes de laudos do Word com segurança e rastreabilidade. O projeto percorre subpastas, ordena os documentos pela **Data do Recebimento**, insere uma sequência no campo **Relatório de Ensaio N°**, valida o resultado, gera uma planilha de controle e converte os DOCX para PDF.

> O repositório não contém documentos reais. Use sempre cópias dos arquivos e confira o lote antes de substituir documentos oficiais.

## Recursos

- busca recursiva de arquivos `.docx`;
- ignora temporários do Word iniciados por `~$`;
- identifica `Data do Recebimento: DD/MM/AAAA` no documento;
- ordena por mês, dia, pasta e nome do arquivo;
- cria números no padrão `00001.2026`;
- substitui um número existente ou preenche um título ainda sem número;
- preserva a estrutura de pastas;
- valida lacunas, duplicidades e alterações fora do título;
- destaca datas cujo ano é diferente do ano do relatório;
- gera `controle_de_laudos.xlsx` e `controle_de_laudos.csv`;
- converte por Microsoft Word no Windows ou LibreOffice;
- interrompe o processamento diante de documento ambíguo ou inválido.

## Requisitos

- Python 3.10 ou superior;
- Microsoft Word no Windows **ou** LibreOffice;
- documentos com os textos `Relatório de Ensaio N°` e `Data do Recebimento: DD/MM/AAAA`.

## Instalação no Windows

```powershell
git clone https://github.com/Victor-Djalma/automacao-laudos-docx-pdf.git
cd automacao-laudos-docx-pdf
py -m venv .venv
.\.venv\Scripts\Activate.ps1
py -m pip install -r requirements.txt
```

Se o PowerShell bloquear a ativação do ambiente virtual, execute uma vez na sessão atual:

```powershell
Set-ExecutionPolicy -Scope Process Bypass
```

## Uso rápido

### 1. Numeração e planilha, sem gerar PDFs

```powershell
py automacao_laudos.py numerar "C:\Laudos\Agosto" "C:\Laudos\Saida" --inicio 1072 --ano 2026
```

A saída será:

```text
Saida/
├── docx_numerados/
├── controle_de_laudos.xlsx
└── controle_de_laudos.csv
```

### 2. Converter DOCX já numerados para PDF

Com seleção automática do motor:

```powershell
py automacao_laudos.py converter "C:\Laudos\Saida\docx_numerados" "C:\Laudos\Saida\pdf"
```

Forçando o Microsoft Word:

```powershell
py automacao_laudos.py converter "C:\Laudos\Saida\docx_numerados" "C:\Laudos\Saida\pdf" --motor word
```

### 3. Executar o fluxo completo

```powershell
py automacao_laudos.py executar "C:\Laudos\Agosto" "C:\Laudos\Saida" --inicio 1072 --ano 2026
```

## Reprocessamento

Por segurança, o programa não sobrescreve uma pasta de saída que já possua arquivos. Para recriá-la deliberadamente, acrescente `--force`:

```powershell
py automacao_laudos.py executar "C:\Laudos\Agosto" "C:\Laudos\Saida" --inicio 1072 --ano 2026 --force
```

## Regras de ordenação

Os laudos são ordenados por:

1. mês da Data do Recebimento;
2. dia;
3. caminho da pasta e nome do arquivo;
4. ano da data, apenas como último critério de desempate.

Essa regra mantém documentos com data de um ano anterior junto ao mesmo dia e mês do lote atual. Esses casos aparecem destacados na aba **Datas fora do ano** da planilha.

## Validações de segurança

Antes de concluir, a automação confere:

- se cada arquivo possui exatamente uma data de recebimento;
- se cada arquivo possui exatamente um título de relatório;
- se todos os números são únicos e contínuos;
- se a árvore de arquivos de entrada e saída corresponde;
- se nenhuma parte interna do DOCX, além do título, foi alterada;
- se o número novo aparece exatamente uma vez no título;
- se cada PDF foi criado e não está vazio.

Essas verificações reduzem erros operacionais, mas não substituem a conferência visual exigida por procedimentos internos ou auditorias.

## Testes

```powershell
py -m pip install -r requirements-dev.txt
py -m pytest -q
```

O GitHub Actions também executa os testes automaticamente a cada atualização do repositório.

## Privacidade

Arquivos `.docx`, `.pdf`, `.xlsx` e pastas comuns de entrada/saída estão bloqueados no `.gitignore`. Assim, documentos reais não são enviados ao GitHub por engano.

## Licença

Distribuído sob a licença [MIT](LICENSE).
