# 📊 Planilhas de Estudo Avançadas

Um repositório dedicado a demonstrar o poder do Microsoft Excel como uma plataforma de aplicação, utilizando técnicas avançadas para criar planilhas de estudo interativas, robustas e com uma experiência de usuário profissional.

## 🚀 Sobre o Projeto

Este projeto vai além das fórmulas tradicionais do Excel. O objetivo é transformar uma simples planilha em uma ferramenta de estudo completa e automatizada. Através do uso de **Formulários (UserForms)** para entrada de dados, **Ribbons Personalizadas** para uma navegação intuitiva, **automação com VBA** e **integração com XML**, criamos uma solução que centraliza e organiza o seu fluxo de estudos de maneira eficiente.

### ✨ Principais Funcionalidades

  * ✍️ **Interface de Cadastro Intuitiva:** Adicione matérias, tópicos e sessões de estudo através de formulários fáceis de usar, evitando a manipulação direta das células.
  * 🧭 **Navegação Profissional:** Uma aba (Ribbon) personalizada no menu superior do Excel agrupa todas as funcionalidades importantes, como "Adicionar Tópico", "Registrar Estudo" e "Gerar Relatório".
  * 🤖 **Automação Inteligente:** Scripts VBA que automatizam tarefas repetitivas, como a organização de dados, cálculos de horas estudadas e a geração de dashboards.
  * ⚙️ **Estrutura Flexível com XML:** A estrutura da Ribbon é definida em XML, permitindo fácil manutenção e expansão da interface do usuário.

## 📸 Screenshots



## 🛠️ Tecnologias e Conceitos Aplicados

Este projeto foi construído utilizando as seguintes tecnologias e conceitos avançados no ambiente Excel:

### 1\. Formulários (UserForms)

Os UserForms são a espinha dorsal da interação do usuário com a planilha.

  - **Entrada de Dados Guiada:** Evita erros e garante que os dados sejam inseridos no formato correto.
  - **Experiência de Aplicação:** O usuário não precisa conhecer Excel a fundo para utilizar a ferramenta, interagindo apenas com botões, caixas de texto e listas.

### 2\. Ribbons Personalizadas (XML/RibbonX)

Para proporcionar uma experiência profissional, a interface padrão do Excel foi estendida com uma Ribbon personalizada.

  - **Custom UI Editor:** Utilizamos a ferramenta *Custom UI Editor for Microsoft Office* para incorporar o código XML que define a estrutura da Ribbon diretamente no arquivo da planilha.
  - **Callbacks VBA:** Cada botão na Ribbon está conectado a uma macro VBA (callback) que executa uma função específica.

### 3\. Automação com VBA (Visual Basic for Applications)

O VBA é o cérebro por trás de toda a automação e lógica da planilha.

  - **Manipulação do DOM (Document Object Model) do Excel:** Leitura e escrita de dados em células, formatação, criação de tabelas e gráficos.
  - **Controle de Eventos:** Execução de código em resposta a ações do usuário, como clicar em um botão ou abrir um formulário.
  - **Lógica de Negócio:** Implementação de regras para cálculos de progresso, validações e gerenciamento do fluxo de dados.

### 4\. Integração com XML

O XML é usado principalmente para a customização da interface do usuário, mas o conceito pode ser expandido para outras funcionalidades.

  - **RibbonX:** Linguagem de marcação baseada em XML usada para definir a estrutura de abas, grupos e botões da Ribbon.
  - **Portabilidade de Dados (Exemplo de uso futuro):** O conhecimento de XML abre portas para importar e exportar dados de e para outras aplicações de forma estruturada.

## ⚙️ Como Utilizar

### Pré-requisitos

  * **Microsoft Excel** (2016 ou superior, preferencialmente Microsoft 365) para Windows.
  * Habilidade para **habilitar a execução de macros**.

### Passos

1.  **Baixe o arquivo:** Faça o download do arquivo `.xlsm` da pasta `/src` (ou da seção de Releases).
2.  **Abra o arquivo no Excel:** Ao abrir, o Excel provavelmente exibirá um aviso de segurança.
3.  **Habilite o Conteúdo:** Clique no botão **"Habilitar Conteúdo"** ou **"Habilitar Macros"** na barra de aviso amarela. Isso é crucial para que o código VBA e a Ribbon personalizada funcionem.
4.  **Explore a Nova Aba:** Após habilitar as macros, uma nova aba com o nome do projeto deverá aparecer no menu superior do Excel. Utilize os botões para interagir com a planilha.

## 🤝 Como Contribuir

Contribuições são o que tornam a comunidade de código aberto um lugar incrível para aprender, inspirar e criar. Qualquer contribuição que você fizer será **muito bem-vinda**.

1.  Faça um **Fork** do projeto.
2.  Crie uma nova **Branch** (`git checkout -b feature/sua-feature-incrivel`).
3.  Faça o **Commit** das suas alterações (`git commit -m 'Adiciona sua-feature-incrivel'`).
4.  Faça o **Push** para a Branch (`git push origin feature/sua-feature-incrivel`).
5.  Abra um **Pull Request**.

## 👤 Autor

  * **Isac Soares Barreto** - *Desenvolvimento e Idealização*
  * **LinkedIn:** `linkedin.com/in/isac-engenheiro-mecatronico/`
  * **GitHub:** `https://github.com/Isac-Soares-Barreto`

## 📜 Licença

Este projeto está licenciado sob a Licença CC BY-NC-SA 4.0. Veja o arquivo `LICENSE` para mais detalhes.
