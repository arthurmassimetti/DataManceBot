# Sistema de Automação de Cadastro de Funcionários - Projeto RH

## Descrição

Este projeto visa desenvolver um sistema de automação para o departamento de Recursos Humanos (RH), focado principalmente em simplificar e agilizar o processo de cadastro de funcionários. O sistema será integrado à base de dados existente da empresa, garantindo a integridade e consistência das informações inseridas. O objetivo é criar uma solução intuitiva e de fácil utilização, permitindo que usuários sem conhecimento técnico avançado possam operá-la com eficiência.

## Objetivos e Metas

- Automatizar o processo de cadastro de funcionários para aumentar a eficiência e reduzir o tempo gasto nesta atividade.
- Integrar-se com a base de dados existente da empresa para acessar e incluir informações pré-existentes dos funcionários.
- Validar os dados inseridos para garantir a integridade e consistência das informações.
- Desenvolver um projeto intuitivo e fácil de usar para permitir a operação do sistema por usuários sem conhecimento técnico avançado.
- Assegurar a segurança dos dados armazenados, garantindo acesso apenas para usuários autorizados.

## Contexto e Justificativa

O departamento de Recursos Humanos desempenha um papel crucial na gestão de pessoas dentro de uma organização. No entanto, o processo manual de cadastro de funcionários pode ser moroso e propenso a erros, resultando em desperdício de tempo e recursos. Automatizar esse processo é fundamental para melhorar a eficiência operacional do departamento de RH, permitindo que os profissionais foquem em atividades mais estratégicas. Além disso, a integração com a base de dados existente ajuda a garantir a consistência das informações e a facilitar a tomada de decisões baseadas em dados. Portanto, o desenvolvimento deste sistema de automação para cadastro de funcionários é essencial para otimizar os processos internos da empresa e promover uma gestão de recursos humanos mais eficaz.

## Instruções de Instalação e Utilização

Para instalar e utilizar este sistema, siga as instruções abaixo:

1. **Python**: Verifique se você possui o Python instalado em sua máquina. Você pode baixá-lo em [python.org](https://www.python.org/downloads/).

2. **Bibliotecas**: Instale as bibliotecas necessárias para o programa. Utilize o comando abaixo para instalar as dependências via pip:

    ```
    pip install pyautogui 
    pip install pandas
    pip install selenium
    ```

3. **Formato da Planilha Excel**:

    Certifique-se de que a planilha Excel esteja formatada corretamente. Siga as instruções abaixo:

    - A primeira linha deve conter o cabeçalho com os nomes das colunas, na seguinte ordem:
        - CPF dos funcionários
        - Nome do funcionário por extenso
        - Tipo de exame
        - Data do exame

4. **Resolução do Monitor**:

    Certifique-se de que seu monitor está configurado para uma resolução de 1440x900 pixels, pois o programa utiliza a biblioteca PyAutoGUI, que depende das coordenadas do monitor.

5. **Executando o Programa**:

    Após as etapas acima serem concluídas com sucesso, você pode executar o programa conforme descrito na seção anterior.
