# 💳 CtrlTransacoesCartao

**CtrlTransacoesCartao** é um sistema desenvolvido em **Visual Basic 6 (VB6)** para o **gerenciamento de transações financeiras com cartão de crédito**, permitindo cadastrar, consultar e exportar relatórios em CSV/Excel.

---

## 🚀 Funcionalidades

- Cadastro e listagem de transações financeiras  
- Exportação de relatórios mensais em formato `.csv`  
- Categorização automática das transações por faixa de valor  
- Integração com banco de dados **SQL Server**  
- Controle de status: `Aprovada`, `Pendente`, `Cancelada`  

---

## 🧩 Estrutura do Banco de Dados

Banco: `dbFinanceiro`  
Tabela principal: `CadastroTransacoes`

| Campo              | Tipo          | Descrição                              	|
|--------------------|---------------|------------------------------------------|
| Id_Transacao       | INT (IDENTITY) | Identificador único da transação      	|
| Numero_Cartao      | VARCHAR(16)   | Número do cartão                       	|
| Valor_Transacao    | DECIMAL(10,2) | Valor da transação                     	|
| Data_Transacao     | DATETIME      | Data e hora da operação                	|
| Descricao          | VARCHAR(255)  | Descrição da transação                 	|
| Status_Transacao   | VARCHAR(20)   | Situação (Aprovada, Pendente, Cancelada) |

---

## 🛠️ Tecnologias Utilizadas

- Visual Basic 6 (VB6)
- SQL Server Express 2022
- ADO (ActiveX Data Objects)
- Windows Forms Clássico

---

## 📦 Instalação e Execução

1. Clone o repositório:
   ```bash
   git clone https://github.com/SEU_USUARIO/CtrlTransacoesCartao.git
