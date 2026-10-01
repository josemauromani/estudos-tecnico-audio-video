# Currículo: Técnico em áudio e vídeo

Vinte e um estudos em quatro grupos. Cada entrada define o nome curto (usado nos arquivos), o título, as
letras e as cores do cabeçalho, a pergunta a que o estudo responde, o que ele cobre e a ideia da planilha.
As pastas e os nomes de arquivo estão em `estudos.txt`.

O escopo de cada estudo é um limite: o que pertence a outro estudo é citado em uma frase, com link, e não
é explicado de novo.

## Grupo 1 · Entender o som e ligar o sistema

### som · Som, frequência e decibel
- Letras: `Hz` (blue, w) `dB` (amber, w)
- Responde a: o que é o som que eu estou controlando, e como ele é medido?
- Cobre: onda sonora, frequência e as faixas (graves, médios, agudos) com instrumentos e voz em cada uma;
  amplitude e volume; decibel como comparação (dB SPL, dBu, dBFS: o que cada um mede); regras práticas
  (+3 dB, +6 dB, +10 dB); lei do inverso do quadrado (dobrar a distância: −6 dB); fase e polaridade;
  sinal analógico e digital em uma página; níveis de exposição e proteção da audição.
- Planilha: calculadora de dB (razão ↔ dB, perda por distância), tabela de faixas de frequência por fonte,
  registro de medição de nível sonoro por ponto da sala.

### cabos · Cabos, conectores e níveis de sinal
- Letras: `XLR` (blue, w) `P10` (amber, w)
- Responde a: que cabo liga o quê, e por que este canal está com ruído?
- Cobre: níveis de sinal (microfone, instrumento, linha, alto-falante) e por que não se misturam;
  balanceado × desbalanceado e como o balanceado cancela ruído; conectores (XLR, P10 TS e TRS, P2, RCA,
  Speakon, combo); direct box (DI) passiva e ativa; phantom power (+48 V): quem precisa e cuidados;
  multicabo e caixa de palco (stage box) analógica e digital; laço de terra e zumbido (ground lift);
  como enrolar, testar e identificar cabos.
- Planilha: inventário de cabos com teste e situação, mapa de ligações (patch) origem → destino,
  tabela de referência conector × sinal.

### mic · Microfones
- Letras: `D` (blue) `C` (amber) `P` (teal)  — dinâmico, condensador, padrão polar
- Responde a: qual microfone usar em cada fonte, e onde colocar?
- Cobre: dinâmico × condensador (× fita, em uma frase); padrões polares (cardioide, supercardioide,
  hipercardioide, omni, figura de 8) e onde cada um rejeita; efeito de proximidade; resposta de frequência;
  formatos (de mão, lapela, headset, gooseneck de púlpito, coral suspenso, overhead, de contato/clip);
  posicionamento por fonte (voz, coral, violão, bateria, amplificador, piano); regra 3:1; técnica de mão
  do cantor/pregador; pop e vento; cuidados e higienização.
- Planilha: inventário de microfones (tipo, padrão, precisa de phantom), escolha de microfone por fonte,
  conferência da regra 3:1 (distâncias).

### caixas · Caixas, amplificação e PA
- Letras: `P` (blue) `A` (amber)
- Responde a: como o som chega ao público com volume e cobertura iguais para todos?
- Cobre: o que é o PA; caixa ativa × passiva; amplificador, potência (RMS, pico) e impedância (ohms,
  ligação em paralelo); sensibilidade e SPL; alto-falantes e vias (woofer, driver, crossover); subwoofer;
  ângulo de cobertura e posicionamento (altura, direção, frente do palco); line array × point source em
  uma página; caixas de delay e front fill; processador/gerenciador; ordem de ligar e desligar;
  proteção (limiter); segurança de montagem (tripés, içamento, energia).
- Planilha: cálculo de impedância em paralelo e casamento amplificador × caixa, estimativa de SPL à
  distância, plano de cobertura por zona da sala.

## Grupo 2 · Operar a mesa e mixar

### mesa · A mesa de som e o caminho do sinal
- Letras: `IN` (blue, w) `MIX` (teal, w) `OUT` (amber, w)
- Responde a: o que cada botão do canal faz, e por onde o sinal passa?
- Cobre: para que serve a mesa; anatomia do canal de cima para baixo (entrada, ganho, phantom, HPF, EQ,
  auxiliares, pan, mute, solo/PFL, fader); barramentos (master LR, grupos/subgrupos, auxiliares, matriz,
  DCA/VCA); pré-fader × pós-fader; inserts; medidores; mesa analógica × digital (camadas, cenas, patch
  digital, controle por tablet); como ler o diagrama de blocos; rotina de zerar a mesa.
  O ajuste de ganho é do estudo `ganho`; EQ é do `eq`; auxiliares de retorno são do `ret`.
- Planilha: folha de canais (input list) com patch, phantom, HPF, destino; mapa de saídas e auxiliares;
  registro de cenas.

### ganho · Ganho e fader: a estrutura de ganho
- Letras: `G` (amber) `F` (teal)
- Vídeo: `AQOI3QxrfV1vMrEH_dlkVhoNcAx8x_GgAk8yhTEXuuWLt2KZTyoD0Tq6I6cTxUJYMhCHPa5UzgNjgo.mp4` (transcrição em `docs/tutoriais/audio/transcricoes.md`). O estudo segue a analogia do vídeo: ganho é a entrada da caixa d'água, fader é a torneira.
- Responde a: quanto ganho dar a cada canal, e por que o fader não conserta um ganho errado?
- Cobre: ganho × fader × volume; o pré-amplificador; ruído de fundo, nível nominal, headroom e clip;
  relação sinal-ruído; medidores (VU analógico com 0 VU, dBFS digital; referência −18 dBFS médio e picos
  até perto de −6 dBFS como ponto de partida); PFL/solo; pad e chave mic/linha; ganho unitário (0 dB do
  fader); estrutura de ganho ao longo da cadeia (canal → grupo → master → amplificador/caixa ativa);
  efeito do ganho nos auxiliares e na microfonia; trim digital.
- Planilha: registro de ganho por canal (fonte, ganho, pico e média lidos, pad, situação calculada),
  cadeia de ganho do sistema, tabela de faixas de referência.

### eq · Equalização
- Letras: `E` (purple) `Q` (blue)
- Responde a: o que cortar e o que realçar para cada som ficar claro?
- Cobre: o que o EQ faz; faixas e como soam (tabela com palavras: embolado, anasalado, sibilante…);
  tipos de filtro (passa-altas/HPF, shelving, bell/paramétrico, semiparamétrico, gráfico);
  frequência, ganho e Q; cortar antes de realçar; método de varredura; HPF por fonte; EQ por fonte
  (voz falada, voz cantada, violão, baixo, bumbo, teclado) como ponto de partida; mascaramento e espaço
  para cada instrumento; EQ de canal × EQ de sistema; EQ gráfico em retorno.
- Planilha: ficha de EQ por canal (HPF e até quatro bandas), tabela de faixas e sintomas, guia de ponto
  de partida por fonte.

### din · Dinâmica e efeitos
- Letras: `C` (red) `G` (slate) `FX` (purple, w)
- Responde a: como controlar o que varia demais, e dar espaço sem embolar?
- Cobre: faixa dinâmica; compressor (threshold, ratio, attack, release, knee, make-up, redução de ganho);
  limiter; gate e expander; de-esser; pontos de partida por fonte; reverb (tipos, tempo, pré-delay) e
  delay (tempo, repetição, tap); envio e retorno de efeito (aux pós-fader); compressão em excesso e
  microfonia; quando não usar.
- Planilha: ficha de dinâmica por canal com cálculo da saída esperada pela razão, ficha de efeitos
  (envios), pontos de partida por fonte.

### ret · Monitores e retornos
- Letras: `AUX` (teal, w) `IEM` (blue, w)
- Responde a: como cada músico se ouve, sem atrapalhar o som da frente?
- Cobre: por que o músico precisa de retorno; caixa de chão (wedge), side fill, fone/in-ear e sistema de
  monitoração pessoal; auxiliar pré-fader; montar uma mixagem de retorno (o que cada um pede); volume de
  palco e o efeito no PA; posição do retorno e padrão polar; EQ de retorno e ring-out (detalhe no `fb`);
  in-ear: vantagens, microfone de ambiente, proteção auditiva; sinais de comunicação músico ↔ operador.
- Planilha: matriz de retornos (canais × mixes, com nível pedido), pedidos por músico, plano de palco
  dos retornos.

### fb · Microfonia
- Letras: `MIC` (red, w) `↻` (slate)
- Vídeo: `AQMMfS8_JQpGs0iAgD1D5V_J1fcfQwkT8GHxzVZ1OwARrg52_BHof53GRgw5aovTFLWMkXUv.mp4` (transcrição em `docs/tutoriais/audio/transcricoes.md`). O estudo segue a analogia do vídeo: a câmera que filma a própria TV.
- Responde a: por que a caixa apita, e como parar o apito sem perder volume?
- Cobre: o ciclo da realimentação; por que apita uma frequência; ganho antes da microfonia; as quatro
  medidas na ordem (posição das caixas, microfone perto da fonte, ganho sem exagero, EQ estreito);
  padrão polar e posição do retorno; número de microfones abertos (cada vez que dobra, −3 dB de margem);
  identificar a frequência (tabela de apito grave/médio/agudo, analisador/RTA); ring-out do sistema passo
  a passo; supressor automático; o que fazer na hora, durante o evento; sala reflexiva.
- Planilha: registro de ring-out (frequência, corte, Q, por saída), diagnóstico de ocorrências
  (quando apitou, canal, causa, medida), margem estimada por microfones abertos.

### semfio · Microfones sem fio
- Letras: `RF` (purple, w)
- Responde a: como evitar falhas, chiados e quedas no microfone sem fio?
- Cobre: transmissor e receptor; faixas (VHF, UHF, 2,4 GHz, digital) e a necessidade de conferir a
  regulamentação local vigente (sem citar número de norma); frequência, canal e grupo; varredura (scan) e
  sincronização; intermodulação e coordenação de vários sistemas; antenas (diversidade, linha de visada,
  distância, distribuidor); squelch e tom piloto; ganho do transmissor; pilhas e baterias; fontes de
  interferência (TV digital, telão de LED, Wi-Fi); rotina antes do evento; plano B com fio.
- Planilha: plano de frequências (sistema, grupo/canal, frequência, usuário) com alerta de frequências
  repetidas ou muito próximas, controle de pilhas/baterias, checklist de RF.

### pass · Passagem de som e mixagem ao vivo
- Letras: `1` (blue) `2` (amber) `3` (teal)
- Responde a: em que ordem montar, passar e mixar, para o evento começar pronto?
- Cobre: preparação (input list, mapa de palco, rider); montagem e line check × soundcheck; ordem da
  passagem (ganho, retornos, PA); construir a mixagem (base, voz por cima, prioridades em culto e em
  evento falado); equilíbrio, panorama e inteligibilidade da palavra; volume da sala (medir, faixa
  combinada); operar durante o evento (mutes, transições, olhar o palco); gravação e mix para a
  transmissão em uma frase (detalhe no `live`); desmontagem e registro do que ajustar.
- Planilha: input list completa, cronograma de passagem, mapa de palco em tabela, registro de
  ocorrências do evento.

## Grupo 3 · Captar a imagem

### expo · Exposição: íris, obturador e ganho
- Letras: `f` (blue) `S` (amber) `ISO` (teal, w)
- Responde a: como deixar a imagem clara na medida, sem estourar nem escurecer?
- Cobre: o triângulo de exposição; abertura (número f, luz e profundidade de campo); obturador
  (velocidade, regra dos 180°, cintilação de lâmpadas e de telão com 60 Hz); ISO/ganho (dB) e ruído;
  filtro ND; ferramentas (zebra, histograma, forma de onda, false color); expor para o rosto em palco
  escuro; faixa dinâmica e contraluz; automático × manual em evento ao vivo; igualar câmeras.
- Planilha: calculadora de stops (equivalência entre f, obturador e ISO), ficha de ajuste por câmera,
  tabela de referência de obturador × taxa de quadros.

### lente · Lentes, foco e cor
- Letras: `mm` (blue, w) `AF` (amber, w) `K` (teal)
- Responde a: que lente usar de onde eu estou, e como manter foco e cor certos?
- Cobre: distância focal e ângulo de visão; tamanho do sensor e fator de corte; zoom óptico × digital;
  lentes fixas × zoom, abertura constante; profundidade de campo; foco manual × automático, peaking,
  ampliação, foco em quem se move; estabilização; balanço de branco e temperatura de cor (Kelvin),
  preset × manual com cartão branco, luz mista; perfis de imagem e igualar cor entre câmeras.
- Planilha: calculadora de enquadramento (largura da cena pela distância, focal e sensor), ficha de
  balanço de branco por câmera, tabela de temperaturas de cor.

### quadro · Enquadramento e movimento de câmera
- Letras: `PG` (blue, w) `PM` (amber, w) `PP` (teal, w)
- Responde a: como compor um quadro que mostra o que importa, e mover a câmera sem distrair?
- Cobre: planos (geral, conjunto, americano, médio, primeiro plano, close, detalhe) e para que serve cada
  um; regra dos terços, espaço acima da cabeça (headroom), espaço do olhar (lead room); linha do
  horizonte; ângulos (normal, plongée, contra-plongée); regra dos 180° (eixo); movimentos (pan, tilt,
  zoom, travelling) e quando usar; tripé: nivelar, cabeça fluida, contrapeso e fricção; câmera na mão e
  estabilizador; posições de câmera em culto e palestra; plano de cobertura com 1, 2 e 3 câmeras;
  comandos do diretor.
- Planilha: lista de planos (shot list) por câmera e momento, plano de posições de câmera, ficha de
  autoavaliação do quadro.

### luz · Iluminação para vídeo e palco
- Letras: `K` (amber) `F` (blue) `C` (teal) — luz principal (key), preenchimento (fill), contraluz
- Responde a: como iluminar para a câmera ver bem, e não só o público?
- Cobre: por que a câmera precisa de mais luz e de luz mais uniforme que o olho; três pontos (principal,
  preenchimento, contraluz) e ângulos; luz dura × suave, difusão e rebatimento; temperatura de cor e IRC
  (CRI); tipos de refletor (LED, fresnel, PAR, elipsoidal, moving head, em uma tabela); intensidade e
  razão de contraste; iluminar púlpito, banda e plateia; fundo e telão de LED/projeção competindo com a
  luz; cintilação (flicker); DMX em uma página; segurança (calor, carga elétrica, fixação com cabo de
  segurança).
- Planilha: plano de luz (refletor, posição, alvo, função, intensidade, temperatura), cálculo de carga
  elétrica por circuito com alerta de sobrecarga, endereçamento DMX.

## Grupo 4 · Produzir e transmitir ao vivo

### sinal · Sinais e cabos de vídeo
- Letras: `HD` (blue, w) `SDI` (amber, w)
- Responde a: por que não aparece imagem, e que cabo e formato usar entre os equipamentos?
- Cobre: resolução (720p, 1080i/p, 4K), taxa de quadros (29,97/30/59,94/60; 25/50) e por que todos os
  equipamentos devem usar o mesmo formato; entrelaçado × progressivo; HDMI × SDI (distância, trava,
  robustez) × NDI/rede × USB; comprimentos típicos e extensores (fibra, HDBaseT); conversores e scalers;
  HDCP e EDID; áudio embutido; proporção de tela; placas de captura; diagnóstico "sem sinal" passo a
  passo; diagrama de sinal do sistema.
- Planilha: mapa de sinais de vídeo (origem, formato, cabo, comprimento, destino) com alerta de formato
  diferente do padrão do sistema e de comprimento acima do recomendado, inventário de cabos e conversores,
  roteiro de diagnóstico.

### corte · Corte ao vivo e direção de imagem
- Letras: `PVW` (teal, w) `PGM` (red, w)
- Responde a: quando cortar, para qual câmera, e como conduzir a equipe?
- Cobre: o switcher (program, preview, corte, mix/fusão, transições, keys e legendas/lower third, PiP);
  multiview; tally e intercom; papéis (diretor, operador de corte, operadores de câmera); linguagem de
  comando ("prepara 2… corta"); motivação do corte (fala, música, reação); ritmo; sequência clássica
  (geral → médio → close); erros de continuidade (salto, eixo); cortar culto: louvor, pregação, avisos,
  leitura; inserção de mídia e de letras; gravação do programa e das câmeras isoladas.
- Planilha: roteiro de corte por momento do evento (câmera principal, alternativas, gráfico), escala
  da equipe com funções, registro de erros para a revisão.

### live · Transmissão ao vivo
- Letras: `LIVE` (red, w)
- Responde a: como levar imagem e som para a internet sem travar, e com áudio bom?
- Cobre: a cadeia (câmeras → switcher → codificador → plataforma → espectador); codificador por software
  (OBS e similares) × hardware; resolução, taxa de quadros, bitrate, codec, keyframe; velocidade de upload
  e folga (bitrate até cerca de metade do upload medido, como regra prática); rede cabeada × Wi-Fi, link
  de reserva; RTMP, chave de transmissão; áudio da transmissão: por que o som da sala não serve, mix
  separado (aux/matriz ou mixagem dedicada), microfone de ambiente, nível e loudness de referência,
  sincronismo áudio × vídeo (delay); cenas e gráficos; monitorar a transmissão; direitos autorais de
  músicas em uma nota; teste antes e roteiro de contingência.
- Planilha: calculadora de bitrate × upload com alerta de folga, ficha de configuração da transmissão,
  checklist de pré-transmissão com horários, registro de quedas.

### proj · Projeção, telão e mídia
- Letras: `LED` (purple, w) `LM` (amber, w)
- Responde a: como mostrar letras, vídeos e slides legíveis para toda a sala?
- Cobre: projetor × painel de LED × TV; lúmens, contraste e luz ambiente; tamanho da tela pela distância
  do último espectador; throw ratio (distância × largura da imagem) e keystone; resolução e proporção;
  software de apresentação (letras, vídeos, Bíblia, avisos): saída de palco e saída principal; tipografia
  legível (linhas por slide, tamanho, contraste); operar letras no tempo da música; formatos de vídeo e
  áudio da mídia; alpha/lower third para a transmissão em uma frase (detalhe no `corte`); cuidados com
  lâmpada, filtro e pixel pitch.
- Planilha: calculadora de projeção (distância, throw ratio, largura; tamanho mínimo pela distância do
  último espectador), lista de mídias do evento em ordem com duração e situação, checklist de slides.

### oper · Operação do evento: roteiro, equipe e imprevistos
- Letras: `A` (blue) `D` (amber) `P` (teal) — antes, durante, depois
- Responde a: quem faz o quê, em que hora, e o que fazer quando algo falha?
- Cobre: o roteiro técnico (run of show) minuto a minuto com deixas de som, vídeo, luz e mídia; funções
  da equipe e comunicação; cronograma de montagem para trás a partir da hora do evento; energia:
  circuitos, carga, aterramento, régua e no-break, ordem de ligar e desligar; segurança (cabos no chão,
  altura, peso, volume); método de diagnóstico (dividir a cadeia ao meio, trocar uma coisa por vez);
  tabela de sintomas → causa provável para som e vídeo; kit de emergência; reunião de revisão depois do
  evento; manutenção e inventário.
- Planilha: roteiro técnico do evento (horário calculado pela duração), escala e cronograma de montagem,
  inventário de equipamentos com manutenção, registro de ocorrências e revisão.
