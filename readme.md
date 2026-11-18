# Sistema Automatizado de Triagem de Currículos 

> Um aplicativo desktop desenvolvido em Python com interface gráfica (Tkinter) para leitura automatizada de currículos em formato PDF, pontuação ponderada baseada em critérios corporativos de Tecnologia da Informação e exportação de relatórios gerenciais em Excel estilizados.

## Sobre o Projeto

Este projeto foi desenvolvido com foco em **engenharia de software defensiva**, **arquitetura limpa (Clean Code)** e **separação de responsabilidades (SoC)**. Ele automatiza o processo de recrutamento e seleção (R&S), varrendo diretórios locais de currículos, avaliando competências técnicas e comportamentais, aplicando notas de corte e gerando uma planilha Excel execututiva com destaque visual para os melhores perfis.

## Principais Funcionalidades

- **Interface Gráfica Intuitiva (Tkinter):** Seleção simplificada de diretórios, configuração de nota de corte e inserção de palavras-chave customizadas.
- **Processamento Concorrente (Threading):** Execução assíncrona em segundo plano para evitar o travamento da interface durante a leitura de grandes lotes de PDFs.
- **Motor de Pontuação Ponderada:** Avalia automaticamente:
  - Anos de experiência no mercado (via expressões regulares - Regex).
  - Formação superior em *Gestão da Tecnologia da Informação (GTI)*.
  - Pós-graduação, MBA ou Mestrado.
  - Senioridade (*Pleno* ou *Sênior*).
  - Proficiência em inglês (*Avançado/Fluente*, *C1/C2*).
  - Volume de cursos tecnológicos e certificações (*AWS, Azure, Scrum, ITIL, SQL, Alura, Udemy*).
  - Palavras-chave/competências extras parametrizáveis pelo usuário.
- **Relatórios Executivos em Excel (`openpyxl` & `pandas`):**
  - Ordenação automática decrescente por pontuação total.
  - Cabeçalho corporativo e efeito zebrado nas linhas.
  - Formatação condicional destacando em verde os candidatos aprovados.
  - Auto-ajuste de largura de colunas.
- **Resiliência e Blindagem (Tratamento de Exceções):**
  - Proteção contra PDFs corrompidos, protegidos por senha ou digitalizados sem texto selecionável.
  - Tratamento de `PermissionError` caso o arquivo Excel de saída esteja aberto durante a exportação.

## Arquitetura do Software

O código segue estritamente os princípios de Orientação a Objetos, dividido em três camadas principais:

1. **`CurriculoProcessor`**: Responsável pela camada de regra de negócio, leitura de arquivos PDF (`pypdf`) e aplicação do motor analítico de pontuação.
2. **`ExcelReportBuilder`**: Camada de persistência de dados e estilização avançada de planilhas (`pandas` e `openpyxl`).
3. **`CurriculoAppUI`**: Camada de apresentação gráfica baseada em eventos, gerenciamento de threads e interações com o usuário (`tkinter`).

## Critérios de Pontuação (Matriz de Avaliação)

| Critério Analisado | Condição Identificada | Pontuação Atribuída |
| :--- | :--- | :--- |
| **Experiência Profissional** | $\ge 5$ anos / $\ge 3$ anos | $+30$ pts / $+15$ pts |
| **Formação Acadêmica** | Gestão da Tecnologia da Informação (GTI) | $+20$ pts |
| **Pós-Graduação** | Pós, MBA ou Mestrado | $+15$ pts |
| **Senioridade** | Sênior / Pleno | $+25$ pts / $+15$ pts |
| **Idioma** | Inglês Avançado / Fluente (ou C1/C2) | $+20$ pts |
| **Portfólio / Cursos** | $\ge 10$ menções / $\ge 5$ menções | $+20$ pts / $+10$ pts |
| **Competências Extras** | Termos customizados inseridos na UI | $+5$ pts por termo |

## Tecnologias Utilizadas

- **Python 3.x**
- **Tkinter** (Interface Gráfica nativa)
- **Pandas** (Manipulação e estruturação de dados tabulares)
- **Openpyxl** (Estilização e manipulação de arquivos Excel `.xlsx`)
- **PyPDF** (Extração de texto de documentos PDF)

## Instalação e Execução

### Pré-requisitos
Certifique-se de ter o **Python** instalado em sua máquina.

1. Instale as dependências necessárias:

pip install pandas openpyxl pypdf

2. Execute a aplicação:

python analisadorCurriculo.py

## Como Usar:

1. Clique em Procurar... na primeira seção para selecionar a pasta contendo os arquivos .pdf dos currículos.

2. No campo de Competências Adicionais, insira tecnologias ou metodologias opcionais separadas por vírgula (ex: python, docker, aws).

3. Defina a Nota de Corte Mínima utilizando o seletor numérico (apenas candidatos com pontuação igual ou superior serão incluídos no relatório).

4. Defina o local e nome do arquivo de saída clicando em Salvar Como.

5. Clique em EXECUTAR ANÁLISE DE CURRÍCULOS e aguarde a mensagem de sucesso!

## Metodologia & Engenharia de Prompt
Este projeto foi concebido e arquitetado através de técnicas avançadas de Engenharia de Prompt, combinando boas práticas de desenvolvimento de software, padrões de projeto (Clean Code e Design Patterns) e automação de processos aplicados à Gestão da Tecnologia da Informação (GTI).

## Autor:
Desenvolvido por Jeremias Silva Santos.