# Agregador de Dados para Declaração de Imposto de Renda no Excel

---

## 🚀 Sobre o Projeto

Este projeto consiste em uma **ferramenta completa desenvolvida no Microsoft Excel**, criada para **organizar e consolidar informações essenciais para a declaração anual de Imposto de Renda**. Com foco na praticidade e usabilidade, a planilha atua como um agregador de dados, permitindo que o usuário registre e controle suas finanças de maneira eficiente, com recursos de navegação intuitiva, validações automáticas para garantir a integridade dos dados e funcionalidades extras como links rápidos para sites relevantes.

---

## ✨ Funcionalidades Principais

* **Menu de Navegação:** Interface amigável com links diretos para as diferentes seções da planilha (Dados Pessoais, Rendimentos, Despesas, Bens e Dívidas).
* **Entrada de Dados Estruturada:** Seções dedicadas para registro de:
    * **Dados Pessoais e Dependentes:** Informações cruciais do contribuinte e seus dependentes.
    * **Rendimentos:** Detalhamento de salários, aluguéis, rendimentos de PJ, etc.
    * **Despesas Dedutíveis:** Registro de gastos com saúde, educação, previdência privada e outros.
    * **Bens e Dívidas:** Controle do patrimônio e obrigações financeiras.
* **Validação Automática de Dados:** Implementação de listas suspensas e regras de validação para minimizar erros de preenchimento e garantir a consistência das informações.
* **Links Rápidos:** Acesso direto a portais importantes como Receita Federal, e-CAC, e outros recursos relevantes para a declaração do IR.
* **Interface Amigável:** Design pensado para ser claro e fácil de usar, mesmo para quem não tem familiaridade avançada com Excel.

---

## 🛠️ Tecnologias Utilizadas

* **Microsoft Excel:** A plataforma principal para a construção e os cálculos da ferramenta.

---

## 📈 Como Utilizar

Para começar a organizar seus dados para o Imposto de Renda:

1.  **Baixe** o arquivo da planilha: `Agregador_IR_Excel.xlsx`.
2.  **Abra-o** no Microsoft Excel.
3.  **Explore o "Menu Principal"** para navegar entre as diferentes seções.
4.  **Preencha as informações** em cada aba (Dados Pessoais, Rendimentos, Despesas, Bens e Dívidas) seguindo as orientações e utilizando as validações de dados para um preenchimento correto.
5.  Utilize os **links rápidos** para acessar informações externas quando necessário.

---

## 📚 Estrutura da Planilha

A planilha está organizada em diversas abas para facilitar a inserção e consulta dos dados. Abaixo, a estrutura de cada aba principal:

### Aba: `_Listas de Apoio` (Oculta)

Esta aba contém as listas usadas para as validações de dados em outras partes da planilha.

| Coluna A                | Coluna B                       | Coluna C                 | Coluna D               |
| :---------------------- | :----------------------------- | :----------------------- | :--------------------- |
| **Tipos de Rendimento** | **Tipos de Despesa Dedutível** | **Tipos de Bens** | **Tipos de Dívidas** |
| `Salário`               | `Médica / Odontológica`        | `Imóvel`                 | `Empréstimo Pessoal`   |
| `Aluguel`               | `Educação`                     | `Veículo`                | `Financiamento Imobiliário` |
| `Autônomo / PJ`         | `Previdência Privada (PGBL)`   | `Conta Corrente / Poupança` | `Financiamento Veículo` |
| `Aposentadoria / Pensão` | `Pensão Alimentícia`           | `Aplicação Financeira (CDB, LCI, LCA)` | `Outras Dívidas`       |
| `Rendimentos de Aplicações Financeiras` | `Livro Caixa`                  | `Ações / Fundos de Investimento` |                        |
| `Outros`                | `Outras`                       | `Outros Bens`            |                        |

### Aba: `Rendimentos`

Registro de todos os rendimentos recebidos durante o ano fiscal.

| Coluna | Descrição da Coluna    | Formato Sugerido | Validação de Dados           |
| :----- | :--------------------- | :--------------- | :--------------------------- |
| `A`    | **Tipo de Rendimento** | Texto            | Lista suspensa da aba `_Listas de Apoio` |
| `B`    | **Fonte Pagadora** | Texto            | N/A                          |
| `C`    | **CNPJ/CPF** | Número           | N/A |
| `D`    | **Valor (R$)** | Moeda            | Somente números decimais     |
| `E`    | **Observações** | Texto            | N/A                          |

### Aba: `Despesas Dedutíveis`

Registro de despesas que podem ser abatidas da base de cálculo do imposto.

| Coluna | Descrição da Coluna    | Formato Sugerido | Validação de Dados           |
| :----- | :--------------------- | :--------------- | :--------------------------- |
| `A`    | **Tipo de Despesa** | Texto            | Lista suspensa da aba `_Listas de Apoio` |
| `B`    | **Prestador** | Texto            | N/A                          |
| `C`    | **CNPJ/CPF** | Número           | N/A |
| `D`    | **Valor (R$)** | Moeda            | Somente números decimais     |
| `E`    | **Data** | Data             | Formato de data válido       |
| `F`    | **Observações** | Texto            | N/A                          |

### Aba: `Bens e Dívidas`

Controle de bens e direitos, bem como de dívidas e ônus reais.

| Coluna | Descrição da Coluna         | Formato Sugerido | Validação de Dados           |
| :----- | :-------------------------- | :--------------- | :--------------------------- |
| `A`    | **Tipo** | Texto            | Lista suspensa da aba `_Listas de Apoio` |
| `B`    | **Descrição Detalhada** | Texto            | N/A                          |
| `C`    | **Local / Endereço / Instituição** | Texto            | N/A                          |
| `D`    | **Valor em 31/12/ANO ANTERIOR (R$)** | Moeda            | Somente números decimais     |
| `E`    | **Valor em 31/12/ANO ATUAL (R$)** | Moeda            | Somente números decimais     |
| `F`    | **Observações** | Texto            | N/A                          |

### Aba: `Dados Pessoais`

Informações cadastrais do declarante e seus dependentes.

| Campo                           | Descrição                              | Formato Sugerido | Validação de Dados                |
| :------------------------------ | :------------------------------------- | :--------------- | :-------------------------------- |
| **Declarante** |                                        |                  |                                   |
| `Nome Completo`                 | Nome completo do contribuinte          | Texto            | N/A                               |
| `CPF`                           | CPF do contribuinte                    | Número           | N/A             |
| `Data de Nascimento`            | Data de nascimento do contribuinte     | Data             | Formato de data válido            |
| `Título de Eleitor`             | Número do título de eleitor            | Número           | N/A                               |
| `Endereço Completo`             | Endereço para declaração               | Texto            | N/A                               |
| `Ocupação Principal`            | Profissão/Ocupação                     | Texto            | N/A                               |
| `Recibo da Última Declaração`   | Número do recibo da declaração anterior | Texto            | N/A (formato específico)          |
| **Dependentes** |                                        |                  |                                   |
| `Nome do Dependente`            | Nome completo do dependente            | Texto            | N/A                               |
| `CPF do Dependente`             | CPF do dependente (se tiver)           | Número           | Exatamente 11 dígitos (Opcional) |
| `Data de Nascimento Dep.`       | Data de nascimento do dependente       | Data             | Formato de data válido            |
| `Grau de Parentesco`            | Ex: Filho(a), Pai/Mãe, Irmão(ã)        | Texto            | N/A (ou lista suspensa)           |

---

## 🧠 Conceitos Aprendidos e Habilidades Desenvolvidas

Durante o desenvolvimento deste projeto, tive a oportunidade de aplicar e aprofundar os seguintes conhecimentos e habilidades:

* **Modelagem de Dados no Excel:** Estruturação lógica de dados em diferentes abas para facilitar a organização e o acesso.
* **Recursos Avançados do Excel:** Utilização intensiva de:
    * **Validação de Dados:** Para criar listas suspensas e regras de entrada de dados, garantindo a qualidade das informações.
    * **Hiperlinks:** Para criar menus de navegação internos e links externos.
    * **Formatação Condicional:** Para destacar informações importantes ou erros.
    * **Proteção de Planilha/Células:** Para proteger fórmulas e garantir a integridade da ferramenta.
    * **Funções Lógicas e de Busca:** (como `SE`, `PROCV`, `SOMA`, etc.) para automatizar cálculos e validações.
* **Design de Interface (UI/UX) no Excel:** Criação de uma interface intuitiva e prática para o usuário, mesmo dentro das limitações do Excel.
* **Documentação Técnica:** A importância de criar um `README.md` claro e conciso, detalhando o propósito do projeto, suas funcionalidades, tecnologias utilizadas e como ele pode ser reproduzido.
* **Controle de Versão com Git e GitHub:** Utilização do Git para gerenciar as versões do projeto e do GitHub para hospedagem e compartilhamento público do repositório, demonstrando práticas de desenvolvimento colaborativo.

---

## 📧 Contato

Se você tiver alguma dúvida, sugestão ou quiser se conectar, sinta-se à vontade para entrar em contato:

* **LinkedIn:** [https://www.linkedin.com/in/guilherme-alves-a576092b1/]
* **Email:** [gmteixeiralves@hotmail.com]

---
