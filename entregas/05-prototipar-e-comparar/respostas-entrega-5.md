# Entrega 5 — respostas da V1

**Arquivos:** [PPTX V1](Entrega_5_Collabora.AI_v1.pptx) e [PDF V1](Entrega_5_Collabora.AI_v1.pdf). Os oito slides mantêm a numeração do caderno (37–44).

**Base:** solução e funções essenciais da [Entrega 4](../04-ideacao-e-decisao/Entrega_4_Collabora.AI.pdf); persona, critérios e códigos de entrevista da [Entrega 3](../03-persona-e-desafio-refinado/Entrega_3_Collabora.AI.pdf).

## Por que estas duas versões

A Aula 8 (slide 28) propõe desenhar mais de um MVP para a mesma ideia, com esforços diferentes, e perguntar qual ensina mais por esforço investido. A Aula 7 (slides 139–154) e o Playbook (Teste da Solução) tratam protótipos de baixa e média fidelidade, de aplicativo e de serviço. Por isso:

- **Versão A** mostra o Caderno Collabora como aplicativo próprio, em três telas de média fidelidade.
- **Versão B** mostra o mesmo caderno sem programar nada, com arquivos comuns (planilha e documento).

As duas representam as duas funções essenciais da Entrega 4, o que permite comparar aspecto por aspecto.

## Slide 38 — planejamento

| Campo | Texto |
| --- | --- |
| Público e situação | Alex e estudantes do 1º e 2º semestres que estudam com colegas. Usam o caderno ao fim de uma sessão de listas, para registrar, e ao voltar a estudar ou na semana da prova, para retomar. |
| Ação ou fluxo principal | Registrar a resolução de um exercício (foto da folha, passos e dica-chave) e, em outra sessão, ver onde o grupo parou e reencontrar um exercício já resolvido. |
| Benefício a comunicar | Reencontrar rápido o que o grupo resolveu e explicou, sem refazer exercícios nem procurar em folhas, arquivos e mensagens (proposta de valor da Entrega 4). |
| Nível de fidelidade | Média na A e baixa na B. Em uma semana, o objetivo é comparar formas de mostrar as duas funções essenciais, não o acabamento visual. |
| Elementos essenciais | Onde o grupo parou; lista de exercícios com situação; registro com foto, passos, dica-chave e dúvida; busca por disciplina, lista ou termo; quem registrou. |
| Hipóteses representadas | Registrar leva menos de dois minutos; foto, passos e dica bastam para refazer o exercício; ver onde parou ajuda a retomar; registro comum evita dúvidas repetidas. |

## Slides 39 e 40 — versões

| Campo | Versão A | Versão B |
| --- | --- | --- |
| Ferramenta | HTML e CSS, renderizados no Google Chrome | LibreOffice Calc e LibreOffice Writer (.xlsx e .docx, abertos também no Google Planilhas e no Google Docs) |
| Tipo | Wireframe de aplicativo de celular, média fidelidade (3 telas) | Protótipo sem código, baixa fidelidade: planilha-índice, ficha-modelo e pastas |
| Demonstra | Retomar pela tela inicial, registrar um exercício e reencontrar uma dica pela busca | A planilha faz o papel da tela inicial; a ficha registra um exercício com foto, passos e dica |
| Decisões | Organização disciplina → lista → exercício; "onde paramos" em destaque; registro pela foto da folha; situação por exercício | Nenhum aplicativo novo; índice e "onde paramos" atualizados à mão; busca só no nome dos arquivos e na planilha |

## Slide 41 — comparação

| Aspecto | Versão A | Versão B | Implicação |
| --- | --- | --- | --- |
| Forma/interface | App com 3 telas; "onde paramos" em destaque | Planilha, ficha e pastas; "onde paramos" é uma linha da planilha | Em A se vê em segundos onde o grupo está |
| Fluxo | Abrir → continuar → registrar em uma tela → buscar | Abrir pasta → copiar ficha → preencher → colar link → atualizar índice | B tem dois passos manuais a mais; o trabalho tende a cair em quem organiza (E2) |
| Funções | Registro, situação, onde paramos, busca no conteúdo | Mesma ficha e situação; onde paramos manual; busca limitada | B perde justamente reencontrar e retomar |
| Tempo e recursos | Programar app ou site e guardar fotos: semanas e algum custo | Uma tarde, com ferramentas gratuitas | B é mais viável hoje |
| Alinhamento | Cumpre os três "deve" da Entrega 3 e o "retomar rápido" | Cumpre registro comum e caminho da resolução; repete o limite de documentos para contas e imagens (E2) | B arrisca virar mais um lugar onde o material se espalha (E4, E5) |

**Diferença decisiva:** retomar e reencontrar sem trabalho manual (função essencial 2 da Entrega 4).

## Slide 42 — decisão

- **Versão escolhida:** A.
- **Motivo:** só a A retoma e reencontra sem passos manuais, o benefício central da proposta de valor; cumpre os três critérios "deve" e responde à frustração da persona Alex. A B é mais barata, mas depende de atualizar o índice à mão.
- **Limitações:** telas estáticas, com dados de exemplo; não testado com estudantes; custo de construir e de guardar fotos incerto; fotos e nomes de estudantes exigem cuidado com a LGPD (Aula 11).

## Pendências antes do envio

- [ ] Preencher a contribuição de cada integrante no slide 44 (está como `TODO`).
- [ ] Conferir a declaração de IA e acrescentar outras ferramentas usadas pela equipe.
- [ ] Se a equipe refizer as telas em outra ferramenta (por exemplo, Figma), trocar o nome da ferramenta nos slides 39 e 42.
- [ ] Exportar como `Entrega_5_Collabora.AI.pdf` e enviar até 12/11 às 23h59.
