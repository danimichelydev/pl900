// Referências aos elementos do DOM
const initialStatusEl = document.getElementById('initial-status');
const loadingMessageEl = document.getElementById('loading-message');
const errorMessageEl = document.getElementById('error-message');

const startAreaEl = document.getElementById('start-area');
const startBtn = document.getElementById('startBtn');

const quizAreaEl = document.getElementById('quiz-area');
const timerEl = document.getElementById('timer');
const questionAreaEl = document.getElementById('question-area');
const questionCounterEl = document.getElementById('question-counter');
const questionEl = document.getElementById('question');
const optionsEl = document.getElementById('options');
const prevBtn = document.getElementById('prevBtn');
const nextBtn = document.getElementById('nextBtn');
const finishBtn = document.getElementById('finishBtn');
const mandatoryMessageEl = document.getElementById('mandatory-message');

const interruptBtn = document.getElementById('interruptBtn'); // Referência ao botão Interromper

const resultsEl = document.getElementById('results');
const scoreEl = document.getElementById('score');
const reviewEl = document.getElementById('review');
const restartBtn = document.getElementById('restartBtn');

// Spans para exibir info na tela de início/resultados
const numQuestionsTextEl = document.getElementById('num-questions-text');
const timerDurationTextEl = document.getElementById('timer-duration-text');
const numQuestionsResultsEl = document.getElementById('num-questions-results');


// Caminho fixo para o arquivo Excel dentro da pasta assets
const EXCEL_FILE_PATH = './assets/perguntas.xlsx';
const NUMBER_OF_QUESTIONS = 20; // Quantidade de questões
const TOTAL_TIMER_SECONDS = 20 * 60; // 20 minutos em segundos

let allQuestionsFromFile = []; // Todas as questões lidas do arquivo
let questions = []; // As 20 questões sorteadas para o simulado atual
let userAnswers = []; // Armazena as respostas do usuário (tamanho 20)

let timerSeconds = TOTAL_TIMER_SECONDS;
let timerIntervalId = null;

// --- Funções do Timer ---
function startTimer() {
    clearInterval(timerIntervalId);
    timerSeconds = TOTAL_TIMER_SECONDS;
    updateTimerDisplay();
    timerIntervalId = setInterval(() => {
        timerSeconds--;
        updateTimerDisplay();
        if (timerSeconds <= 0) {
            clearInterval(timerIntervalId);
            timerEl.textContent = "00:00";
            // Quando o tempo acaba, finaliza o quiz (sem exigir resposta obrigatória)
            finishQuiz(false); // Passa false para indicar que não foi clique manual (nem no Finish, nem no Interrupt)
        }
    }, 1000);
}

function updateTimerDisplay() {
    const minutes = Math.floor(timerSeconds / 60);
    const seconds = timerSeconds % 60;
    const formattedTime = `${String(minutes).padStart(2, '0')}:${String(seconds).padStart(2, '0')}`;
    timerEl.textContent = formattedTime;
     if (timerSeconds <= 60) { // último minuto
        timerEl.style.color = '#d9534f';
    } else {
        timerEl.style.color = '#d9534f';
    }
}

function stopTimer() {
    clearInterval(timerIntervalId);
    timerIntervalId = null;
}

// --- Funções de Carregamento Automático do Arquivo Excel ---

// Função para buscar o arquivo Excel
async function fetchExcelFile(url) {
    initialStatusEl.style.display = 'block';
    loadingMessageEl.style.display = 'block';
    errorMessageEl.style.display = 'none';
    startAreaEl.style.display = 'none';
    quizAreaEl.style.display = 'none';
    resultsEl.style.display = 'none';

    try {
        const response = await fetch(url);
        if (!response.ok) {
            throw new Error(`Erro HTTP: ${response.status} ${response.statusText}. Verifique se o arquivo "${url}" existe e está acessível no servidor.`);
        }
        const arrayBuffer = await response.arrayBuffer();
        const workbook = XLSX.read(arrayBuffer, { type: 'array' });
        const firstSheetName = workbook.SheetNames[0];
        const worksheet = workbook.Sheets[firstSheetName];
        // raw:true tenta preservar tipos de dados originais (útil para comentários que podem ser números, datas, etc)
        // defval:'' define um valor padrão para células vazias, evitando 'undefined' ou 'null' inesperados
        const jsonData = XLSX.utils.sheet_to_json(worksheet, { header: 1, raw: true, defval:'' });

        allQuestionsFromFile = processParsedData(jsonData);

        loadingMessageEl.style.display = 'none';

        if (allQuestionsFromFile.length >= NUMBER_OF_QUESTIONS) {
            console.log(`Carregadas ${allQuestionsFromFile.length} questões válidas do arquivo Excel (incluindo comentários).`);
            // Atualiza os spans na tela de início com a contagem e duração
            numQuestionsTextEl.textContent = NUMBER_OF_QUESTIONS;
            timerDurationTextEl.textContent = TOTAL_TIMER_SECONDS / 60; // Mostra em minutos
            numQuestionsResultsEl.textContent = NUMBER_OF_QUESTIONS; // Atualiza na tela de resultados tbm

            initialStatusEl.style.display = 'none';
            startAreaEl.style.display = 'block'; // Mostra a área de início
        } else if (allQuestionsFromFile.length > 0) {
             errorMessageEl.textContent = `O arquivo Excel contém apenas ${allQuestionsFromFile.length} questões válidas. É necessário pelo menos ${NUMBER_OF_QUESTIONS} para iniciar o simulado.`;
             errorMessageEl.style.display = 'block';
        }
        else {
             errorMessageEl.textContent = `Nenhuma questão válida encontrada no arquivo Excel. Verifique a estrutura das colunas (A-G) e se há pelo menos ${NUMBER_OF_QUESTIONS} questões.`;
             errorMessageEl.style.display = 'block';
        }

    } catch (error) {
        console.error("Erro ao buscar ou processar o arquivo Excel:", error);
        loadingMessageEl.style.display = 'none';
        errorMessageEl.textContent = `Não foi possível carregar as questões: ${error.message}. Verifique se o arquivo "${EXCEL_FILE_PATH}" existe e está acessível no servidor.`;
        errorMessageEl.style.display = 'block';
    }
}

// Processa o array de dados lido do arquivo para o formato de perguntas
function processParsedData(jsonData) {
    const loadedQuestions = [];
    // Itera a partir da segunda linha (índice 1) para pular o cabeçalho
    for (let i = 1; i < jsonData.length; i++) {
        const row = jsonData[i];
        // Verifica se a linha tem colunas suficientes (AGORA pelo menos 7: A-G)
        // E não é apenas uma linha completamente vazia ou com espaços
         if (row && row.length >= 7 && row.some(cell => cell !== undefined && cell !== null && String(cell).trim() !== '')) {
             const questionText = row[0];
             const optionA = row[1];
             const optionB = row[2];
             const optionC = row[3];
             const optionD = row[4];
             const correctAnswer = row[5] ? String(row[5]).toUpperCase().trim() : null;
             const commentText = row[6] ? String(row[6]).trim() : ''; // Le a coluna G (indice 6), padrao '' se vazia

             // Validação básica: Pergunta e opções A-D não devem ser vazias/nulas após trim, e a resposta deve ser A, B, C ou D.
             if (
                 questionText && String(questionText).trim() !== '' &&
                 optionA && String(optionA).trim() !== '' &&
                 optionB && String(optionB).trim() !== '' &&
                 optionC && String(optionC).trim() !== '' &&
                 optionD && String(optionD).trim() !== '' &&
                 ['A', 'B', 'C', 'D'].includes(correctAnswer)
                 ) {
                loadedQuestions.push({
                    question: String(questionText).trim(),
                    options: {
                        A: String(optionA).trim(),
                        B: String(optionB).trim(),
                        C: String(optionC).trim(),
                        D: String(optionD).trim()
                    },
                    correctAnswer: correctAnswer,
                    comment: commentText // Armazena o comentario
                });
             } else {
                 // console.warn(`Linha ${i+1} ignorada devido a dados inválidos ou incompletos:`, row);
             }
        } else {
             // console.warn(`Linha ${i+1} ignorada por não ter colunas suficientes ou estar vazia:`, row);
        }
    }
    return loadedQuestions;
}

// Função para embaralhar um array (Fisher-Yates Shuffle) (Sem mudanças)
function shuffleArray(array) {
    const shuffled = [...array];
    for (let i = shuffled.length - 1; i > 0; i--) {
        const j = Math.floor(Math.random() * (i + 1));
        [shuffled[i], shuffled[j]] = [shuffled[j], shuffled[i]];
    }
    return shuffled;
}

// --- Funções de Controle do Fluxo do Quiz ---

// Prepara as questões sorteadas e inicializa as respostas
function prepareQuizQuestions() {
     if (allQuestionsFromFile.length < NUMBER_OF_QUESTIONS) {
         errorMessageEl.textContent = `Erro: Não há questões suficientes para iniciar o simulado. (${allQuestionsFromFile.length} encontradas, ${NUMBER_OF_QUESTIONS} necessárias).`;
         errorMessageEl.style.display = 'block';
         initialStatusEl.style.display = 'block';
         startAreaEl.style.display = 'none';
         quizAreaEl.style.display = 'none';
         resultsEl.style.display = 'none';
         return false;
     }

    const shuffledAllQuestions = shuffleArray(allQuestionsFromFile);
    questions = shuffledAllQuestions.slice(0, NUMBER_OF_QUESTIONS); // Pega as 20 questões

    currentQuestionIndex = 0;
    userAnswers = Array(NUMBER_OF_QUESTIONS).fill(null); // Inicializa respostas

    return true;
}

// Inicia a exibição do quiz (mostra área, inicia timer, carrega primeira pergunta)
function displayQuiz() {
    startAreaEl.style.display = 'none';
    resultsEl.style.display = 'none';
    quizAreaEl.style.display = 'block'; // Mostra o wrapper do quiz
    questionAreaEl.style.display = 'block';
    mandatoryMessageEl.style.display = 'none';
    interruptBtn.style.display = 'block'; // Mostra o botão Interromper

    startTimer(); // Inicia o temporizador!
    loadQuestion(currentQuestionIndex); // Carrega a primeira pergunta
}

// Handler para o botão "Iniciar Simulado"
function handleStartQuiz() {
     if (prepareQuizQuestions()) { // Prepara as perguntas. Se bem-sucedido:
         displayQuiz(); // Inicia a exibição do quiz
     }
}

// Handler para o botão "Interromper Simulado"
function handleInterruptQuiz() {
     const confirmInterrupt = confirm("Tem certeza que deseja interromper o simulado e ver os resultados?");
     if (confirmInterrupt) {
         saveAnswer(); // Salva a resposta atual antes de finalizar
         finishQuiz(false); // Finaliza o quiz (passa false para não exigir resposta obrigatória)
     }
}


// --- Funções do Quiz (Sem mudanças significativas na lógica do quiz) ---

// Função para carregar e exibir uma pergunta (Sem mudanças)
function loadQuestion(index) {
    if (questions.length === 0 || index < 0 || index >= questions.length) {
        console.error("Erro interno: Índice de pergunta inválido ou sem perguntas sorteadas.");
        stopTimer();
        initialStatusEl.style.display = 'block';
        errorMessageEl.textContent = "Ocorreu um erro interno ao carregar a pergunta. Por favor, recarregue a página.";
        errorMessageEl.style.display = 'block';
        quizAreaEl.style.display = 'none';
        resultsEl.style.display = 'none';
        return;
    }
    const questionData = questions[index];
    questionCounterEl.textContent = `Questão ${index + 1} de ${questions.length}`;
    questionEl.textContent = questionData.question;
    optionsEl.innerHTML = '';

    mandatoryMessageEl.style.display = 'none';

    if (questionData.options) {
        const optionKeys = ['A', 'B', 'C', 'D'];
        optionKeys.forEach(optionKey => {
             const optionValue = questionData.options[optionKey];
             const optionId = `option-${index}-${optionKey}`;
             const optionName = `answer-${index}`;

            if (optionValue !== undefined && optionValue !== null && String(optionValue).trim() !== '') {
                const optionWrapper = document.createElement('div');
                const input = document.createElement('input');
                input.type = 'radio';
                input.name = optionName;
                input.id = optionId;
                input.value = optionKey;

                const label = document.createElement('label');
                label.htmlFor = optionId;
                label.textContent = `${optionKey}) ${optionValue}`;

                if (userAnswers[index] === optionKey) {
                    input.checked = true;
                }

                optionWrapper.appendChild(input);
                optionWrapper.appendChild(label);
                optionsEl.appendChild(optionWrapper);
            }
        });

        // Adiciona evento change à div optionsEl apenas uma vez (delegação)
         if (!optionsEl.dataset.listenerAdded) {
             optionsEl.addEventListener('change', handleOptionChange);
             optionsEl.dataset.listenerAdded = 'true';
         }

    } else {
        console.error("Opções ausentes ou inválidas para a questão:", questionData);
        optionsEl.innerHTML = "<p>Erro ao carregar opções para esta questão.</p>";
        prevBtn.disabled = true;
        nextBtn.disabled = true;
        finishBtn.style.display = 'none';
        stopTimer();
        mandatoryMessageEl.textContent = "Erro ao carregar opções desta questão.";
        mandatoryMessageEl.style.color = 'red';
        mandatoryMessageEl.style.display = 'block';
        return;
    }

    prevBtn.disabled = (currentQuestionIndex === 0);

    if (currentQuestionIndex === questions.length - 1) {
        nextBtn.style.display = 'none';
        finishBtn.style.display = 'inline-block';
    } else {
        nextBtn.style.display = 'inline-block';
        finishBtn.style.display = 'none';
    }
}

// Handler para o evento change na div optionsEl (delegação) (Sem mudanças)
function handleOptionChange(event) {
    if (event.target.type === 'radio' && event.target.name === `answer-${currentQuestionIndex}`) {
        mandatoryMessageEl.style.display = 'none';
    }
}

// Função para salvar A resposta selecionada pelo usuário para A pergunta atual (Sem mudanças)
function saveAnswer() {
    const selectedOption = optionsEl.querySelector(`input[name="answer-${currentQuestionIndex}"]:checked`);
    if (selectedOption) {
        userAnswers[currentQuestionIndex] = selectedOption.value;
         return true;
    } else {
        userAnswers[currentQuestionIndex] = null;
        return false;
    }
}

// Navegar para a próxima pergunta (Sem mudanças na lógica de obrigatoriedade)
function nextQuestion() {
    const isAnswerSelected = optionsEl.querySelector(`input[name="answer-${currentQuestionIndex}"]:checked`) !== null;
    if (!isAnswerSelected) {
        mandatoryMessageEl.textContent = "Por favor, selecione uma opção para continuar.";
        mandatoryMessageEl.style.color = 'red';
        mandatoryMessageEl.style.display = 'block';
        return;
    }
    saveAnswer();
    mandatoryMessageEl.style.display = 'none';
    if (currentQuestionIndex < questions.length - 1) {
        currentQuestionIndex++;
        loadQuestion(currentQuestionIndex);
    }
}

// Navegar para a pergunta anterior (Sem mudanças)
function prevQuestion() {
    saveAnswer();
    mandatoryMessageEl.style.display = 'none';
    if (currentQuestionIndex > 0) {
        currentQuestionIndex--;
        loadQuestion(currentQuestionIndex);
    }
}

// Finalizar o simulado e exibir resultados
// Recebe um parâmetro opcional `isManualFinish` (true se clicou em Finalizar, false se timer ou Interromper)
function finishQuiz(isManualFinish = true) { // Default é true para cliques no botão Finalizar
     // Verifica se UMA resposta foi SELECIONADA para a última pergunta SOMENTE se clicou no botão FINALIZAR
     if (isManualFinish && currentQuestionIndex === questions.length - 1) {
        const isAnswerSelected = optionsEl.querySelector(`input[name="answer-${currentQuestionIndex}"]:checked`) !== null;
         if (!isAnswerSelected) {
            mandatoryMessageEl.textContent = "Por favor, selecione uma opção para finalizar o simulado.";
            mandatoryMessageEl.style.color = 'red';
            mandatoryMessageEl.style.display = 'block';
            return; // Sai da função sem finalizar
         }
          saveAnswer(); // Salva a resposta da última pergunta se foi selecionada
     } else {
         // Se finishQuiz foi chamado pelo timer ou Interromper, apenas salva o estado atual da pergunta
          saveAnswer();
     }

    stopTimer(); // Para o timer

    let score = 0;
    reviewEl.innerHTML = ''; // Limpa a revisão anterior

    // Calcula a pontuação e prepara a revisão
    questions.forEach((question, index) => {
        const userAnswer = userAnswers[index];
        const correctAnswer = question.correctAnswer;
        // Considera correta apenas se a resposta do usuário não for null E for igual à correta
        const isCorrect = (userAnswer !== null && userAnswer === correctAnswer);


        if (isCorrect) {
            score++;
        }

        const reviewItem = document.createElement('div');
        reviewItem.classList.add('question-review');

         let userAnswerText = 'Não respondido'; // Texto padrão se não respondeu
         let userAnswerValue = null; // Para aplicar classe de cor
         if (userAnswer !== null && question.options && question.options[userAnswer]) {
             userAnswerText = `${userAnswer}) ${question.options[userAnswer]}`;
             userAnswerValue = userAnswer; // Define para aplicar cor se respondeu
         } else if (userAnswer !== null) { // Caso a resposta esteja salva (ex: 'A') mas a opção não exista mais
              userAnswerText = `${userAnswer}) Opção não encontrada no arquivo original`;
              userAnswerValue = userAnswer;
         }

         let correctAnswerText = 'N/A';
          if (correctAnswer && question.options && question.options[correctAnswer]) {
             correctAnswerText = `${correctAnswer}) ${question.options[correctAnswer]}`;
         } else if (correctAnswer) {
              correctAnswerText = `${correctAnswer}) Opção não encontrada no arquivo original`;
         }

        reviewItem.innerHTML = `
            <p><strong>Questão ${index + 1}:</strong> ${question.question}</p>
            <p>Sua Resposta: <span class="user-answer">${userAnswerText}</span></p>
            <p>Resposta Correta: <span class="correct-answer">${correctAnswerText}</span></p>
            <p class="feedback ${isCorrect ? 'correct-feedback' : (userAnswer !== null ? 'incorrect-feedback' : '')}">${isCorrect ? 'Correto!' : (userAnswer !== null ? 'Incorreto!' : 'Não respondida')}</p>
        `;

         // Adiciona o comentário SE ele existir para esta questão
         if (question.comment && String(question.comment).trim() !== '') {
             const commentEl = document.createElement('p');
             commentEl.classList.add('explanation');
             commentEl.textContent = question.comment; // Usa textContent para segurança e preservar quebras de linha básicas
             reviewItem.appendChild(commentEl);
         }


         // Aplica a classe de cor baseada se respondeu, e se acertou/errou
         const userAnswerSpan = reviewItem.querySelector('.user-answer');
         if (userAnswerValue !== null) { // Aplica cor apenas se o usuário respondeu algo
            userAnswerSpan.classList.add(isCorrect ? 'correct-feedback' : 'incorrect-feedback');
            if (!isCorrect) { // Se respondeu, mas errou
                 reviewItem.querySelector('.correct-answer').style.fontWeight = 'bold';
            }
         } else { // Se não respondeu
             userAnswerSpan.style.fontStyle = 'italic'; // Itálico para "Não respondido"
             reviewItem.querySelector('.correct-answer').style.fontWeight = 'bold'; // Destaca a correta
         }


        reviewEl.appendChild(reviewItem);
    });

    const totalQuestions = questions.length;
    const percentage = totalQuestions > 0 ? ((score / totalQuestions) * 100).toFixed(2) : 0;
    scoreEl.textContent = `Você acertou ${score} de ${totalQuestions} perguntas. (${percentage}%)`;

    // Esconde a área do quiz e mostra a área de resultados
    quizAreaEl.style.display = 'none';
    resultsEl.style.display = 'block';
    interruptBtn.style.display = 'none'; // Garante que o botão Interromper está oculto nos resultados
}

// Reiniciar o simulado (sorteia novas 20 questões do total carregado e inicia o fluxo do quiz)
function restartQuiz() {
     if (allQuestionsFromFile.length < NUMBER_OF_QUESTIONS) {
         stopTimer();
         initialStatusEl.style.display = 'block';
         errorMessageEl.textContent = "Não há questões suficientes para reiniciar o simulado. Recarregue a página para tentar carregar o arquivo novamente.";
         errorMessageEl.style.display = 'block';
         startAreaEl.style.display = 'none';
         quizAreaEl.style.display = 'none';
         resultsEl.style.display = 'none';
         return;
     }
    if (prepareQuizQuestions()) {
        displayQuiz(); // Inicia a exibição do quiz (inclui iniciar novo timer)
    }
}

// Recarregar a página para tentar carregar o arquivo novamente
// Mantida, embora os botões diretos tenham sido removidos.
function reloadPage() {
    window.location.reload();
}


// --- Adiciona Ouvintes de Evento ---
startBtn.addEventListener('click', handleStartQuiz); // Listener para o botão Iniciar
interruptBtn.addEventListener('click', handleInterruptQuiz); // Listener para o botão Interromper

prevBtn.addEventListener('click', prevQuestion);
// Ao clicar em Próxima ou Finalizar, a verificação de obrigatoriedade é feita DENTRO dessas funções.
nextBtn.addEventListener('click', nextQuestion);
finishBtn.addEventListener('click', () => finishQuiz(true)); // Passa true para indicar clique no Finalizar

restartBtn.addEventListener('click', restartQuiz);

// Listener único para a div optionsEl usando delegação para capturar clicks em radio buttons
// Adicionado apenas uma vez no carregamento inicial do script
if (!optionsEl.dataset.listenerAdded) {
    optionsEl.addEventListener('change', handleOptionChange);
    optionsEl.dataset.listenerAdded = 'true';
}


// --- Inicialização ---
// Ao carregar a página, tenta buscar o arquivo Excel automaticamente
fetchExcelFile(EXCEL_FILE_PATH);