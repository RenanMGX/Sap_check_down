# SAP Check Automation

Esse repositório contém um script automatizado para monitorar e interagir com a janela do SAP GUI. O script verifica se uma janela com o título "SAP GUI for Windows 770" está ativa. Se estiver, ele a traz para o primeiro plano e simula o pressionamento da tecla Enter.

## Estrutura do Projeto

- **main.py**: Script principal que realiza a verificação e interação com a janela do SAP.
- **Entities/sap_check.py**: Contém a classe `SapCheck` com métodos para:
  - Buscar a janela desejada (`find_per_title`);
  - Trazer a janela para o primeiro plano (`para_frente`);
  - Simular o pressionamento da tecla Enter (`aperta_enter`).
- **Entities/dependencies/functions.py**: Provavelmente contém funções auxiliares, como a função `P` para formatação de mensagens no console.
- **requirements.txt**: Lista as dependências do projeto, incluindo a biblioteca `pywin32` para interação com a API do Windows.

## Fluxo do Código em `main.py`

```python
if __name__ == "__main__":
    count = 1
    max_range = 10
    for i in range(max_range):
        try:
            sap = SapCheck()
            # Busca a janela que contenha o texto "SAP GUI for Windows 770"
            janela = sap.find_per_title("SAP GUI for Windows 770")
            if janela:
                # Traz a janela para o primeiro plano
                sap.para_frente(janela[0])
                # Simula o pressionamento da tecla Enter na janela
                sap.aperta_enter(janela[0])
                print(P(f"[{count}/{max_range}] A janela {janela} foi fechada com sucesso!", color="green"))
            else:
                print(P(f"[{count}/{max_range}]", color="red"))
        except Exception as err:
            print(P(f"[{count}/{max_range}] - {err}", color="red"))
            print(traceback.format_exc())
        
        count += 1
        del sap
        sleep(1)
```

### Descrição do Fluxo:

1. **Inicialização**:  
   O script inicia definindo um contador (`count`) e um número máximo de iterações (`max_range`).

2. **Loop de Processamento**:  
   Para cada iteração:
   - Cria uma instância da classe `SapCheck`.
   - Tenta encontrar uma janela cujo título contenha "SAP GUI for Windows 770" usando o método `find_per_title`.
   - Se a janela for encontrada:
     - Ela é trazida para o primeiro plano com o método `para_frente`.
     - O script simula o pressionamento da tecla Enter através do método `aperta_enter`.
     - Uma mensagem de sucesso é exibida.
   - Se a janela não for encontrada, uma mensagem com status de erro (ou ausência) é exibida.
   - Em caso de exceção, o erro e o traceback são impressos.

3. **Limpeza e Controle de Fluxo**:  
   - O objeto `sap` é deletado ao final de cada iteração.
   - O script aguarda 1 segundo (`sleep(1)`) antes de iniciar a nova iteração.

## Dependências e Execução

- **Dependências**:  
  As dependências necessárias estão listadas no arquivo `requirements.txt`. Certifique-se de que a biblioteca `pywin32` (entre outras) esteja instalada.

- **Como Executar**:  
  Para executar o script, utilize o seguinte comando no terminal:

  ```cmd
  python main.py
  ```

## Observações

- O método de trazer a janela para o primeiro plano pode estar sujeito a restrições do sistema operacional, que podem impedir alterações de foco vindas de processos em segundo plano.
- A simulação do pressionamento da tecla Enter é feita utilizando funções da biblioteca `pywin32` para replicar eventos de teclado.
