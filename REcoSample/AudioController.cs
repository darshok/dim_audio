using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.IO;
using System.Text;
using System.Windows.Forms;
using System.Speech.Recognition;
using System.Speech.Synthesis;

namespace SpeechDIM
{
    public partial class AudioController : Form
    {
        private SpeechRecognitionEngine recognizer =
           new SpeechRecognitionEngine();
        private SpeechSynthesizer synth = new SpeechSynthesizer();

        public AudioController()
        {
            InitializeComponent();
        }


        private void AudioController_Load(object sender, EventArgs e)
        {
            synth.Speak("Inicializando la Aplicación");

            Grammar grammar = CreateAudioControllerGrammar();
            recognizer.SetInputToDefaultAudioDevice();
            recognizer.UnloadAllGrammars();
            recognizer.UpdateRecognizerSetting("CFGConfidenceRejectionThreshold", 60);
            grammar.Enabled = true;
            recognizer.LoadGrammar(grammar);
            recognizer.SpeechRecognized += new EventHandler<SpeechRecognizedEventArgs>(recognizer_SpeechRecognized);
            //reconocimiento asíncrono y múltiples veces
            recognizer.RecognizeAsync(RecognizeMode.Multiple);
            synth.Speak("Aplicación preparada para reconocer su voz");
        }


        void recognizer_SpeechRecognized(object sender, SpeechRecognizedEventArgs e)
        {
            //obtenemos un diccionario con los elementos semánticos
            SemanticValue semantics = e.Result.Semantics;

            string rawText = e.Result.Text;
            RecognitionResult result = e.Result;

            if (!semantics.ContainsKey("rgb"))
            {
                label1.Text = rawText;
                synth.Speak(rawText);
            }
            else
            {
                label1.Text = rawText;
                BackColor = Color.FromArgb((int)semantics["rgb"].Value);
                Update();
                synth.Speak(rawText);
            }
        }

        private Grammar CreateAudioControllerGrammar()
        {
            synth.Speak("Creando ahora la gramática");
            Choices colorChoice = new Choices();

            // foreach (string colorName in System.Enum.GetNames(typeof(KnownColor)))
            // {
            //     SemanticResultValue choiceResultValue =
            //         new SemanticResultValue(colorName, Color.FromName(colorName).ToArgb());
            //     GrammarBuilder resultValueBuilder = new GrammarBuilder(choiceResultValue);
            //     colorChoice.Add(resultValueBuilder);
            // }

            SemanticResultValue choiceResultValue =
                    new SemanticResultValue("Rojo", Color.FromName("Red").ToArgb());
            GrammarBuilder resultValueBuilder = new GrammarBuilder(choiceResultValue);
            colorChoice.Add(resultValueBuilder);

            choiceResultValue =
                   new SemanticResultValue("Azul", Color.FromName("Blue").ToArgb());
            resultValueBuilder = new GrammarBuilder(choiceResultValue);
            colorChoice.Add(resultValueBuilder);

            choiceResultValue =
                   new SemanticResultValue("Verde", Color.FromName("Green").ToArgb());
            resultValueBuilder = new GrammarBuilder(choiceResultValue);
            colorChoice.Add(resultValueBuilder);

            SemanticResultKey choiceResultKey = new SemanticResultKey("rgb", colorChoice);
            GrammarBuilder colores = new GrammarBuilder(choiceResultKey);

            /*
             * Pausar/parar la canción
             * Reproducir/continuar la canción
             * frase()
             * 
             * Aumentar/subir el volumen
             * Reducir/bajar el volumen
             * frase()
             * 
             * Avanzar/Saltar a la siguiente canción
             * Volver/Retroceder a la anterior canción
             * frase()
             * 
             * como añadir mas frases a una grammar
             * https://docs.microsoft.com/en-us/dotnet/api/system.speech.recognition.grammarbuilder?view=netframework-4.8
             */

            GrammarBuilder la = "la";
            GrammarBuilder el = "el";
            GrammarBuilder a = "a";
            GrammarBuilder cancion = "cancion";
            GrammarBuilder pausar = "Pausar";
            GrammarBuilder parar = "Parar";
            GrammarBuilder reproducir = "Reproducir";
            GrammarBuilder continuar = "Continuar";

            Choices resumePauseChoices = new Choices(pausar, parar, reproducir, continuar); //habrá que hacerlo con los semanticresultvalue
            GrammarBuilder phraseResumePause = new GrammarBuilder(resumePauseChoices);
            phraseResumePause.Append(la);
            phraseResumePause.Append(cancion);

            GrammarBuilder aumentar = "Aumentar";
            GrammarBuilder subir = "Subir";
            GrammarBuilder reducir = "Reducir";
            GrammarBuilder bajar = "Bajar";
            GrammarBuilder volumen = "volumen";

            Choices volumeChoices = new Choices(aumentar, subir, reducir, bajar); //habrá que hacerlo con los semanticresultvalue
            GrammarBuilder phraseVolume = new GrammarBuilder(volumeChoices);
            phraseVolume.Append(el);
            phraseVolume.Append(volumen);


            GrammarBuilder avanzar = "Avanzar";
            GrammarBuilder saltar = "Saltar";
            GrammarBuilder volver = "Volver";
            GrammarBuilder retroceder = "Retroceder";
            GrammarBuilder siguiente = "siguiente";
            GrammarBuilder anterior = "anterior";

            Choices controlChoices = new Choices(avanzar, saltar, volver, retroceder); //habrá que hacerlo con los semanticresultvalue
            GrammarBuilder phraseControl = new GrammarBuilder(controlChoices);
            phraseControl.Append(a);
            phraseControl.Append(la);
            Choices laElChoices = new Choices(siguiente, anterior);
            phraseControl.Append(laElChoices);
            phraseControl.Append(cancion);

            GrammarBuilder poner = "Poner";
            GrammarBuilder cambiar = "Cambiar";
            GrammarBuilder fondo = "Fondo";

            Choices colorChoices = new Choices(poner, cambiar);
            GrammarBuilder phraseColors = new GrammarBuilder(colorChoices);
            phraseColors.Append(fondo);
            phraseColors.Append(colores);

            Choices allPhrasesChoices = new Choices(new GrammarBuilder[] { phraseResumePause, phraseVolume, phraseControl, phraseColors });

            Grammar grammar = new Grammar((GrammarBuilder)allPhrasesChoices);
            grammar.Name = "Audio controller";
            return grammar;
        }

    }
}