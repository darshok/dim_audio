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
using System.Runtime.InteropServices;

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

        /*
         * VK_VOLUME_MUTE 0xAD Volume Mute key
         * VK_VOLUME_DOWN 0xAE Volume Down key
         * VK_VOLUME_UP 0xAF Volume Up key
         * VK_MEDIA_NEXT_TRACK 0xB0 Next Track key
         * VK_MEDIA_PREV_TRACK 0xB1 Previous Track key
         * VK_MEDIA_STOP 0xB2 Stop Media key
         * VK_MEDIA_PLAY_PAUSE 0xB3 Play/Pause Media key
         */

        private const byte VK_VOLUME_MUTE = 0xAD;
        private const byte VK_VOLUME_DOWN = 0xAE;
        private const byte VK_VOLUME_UP = 0xAF;
        private const byte VK_MEDIA_NEXT_TRACK = 0xB0;
        private const byte VK_MEDIA_PREV_TRACK = 0xB1;
        private const byte VK_MEDIA_STOP = 0xB2;
        private const byte VK_MEDIA_PLAY_PAUSE = 0xB3;
        private const UInt32 KEYEVENTF_EXTENDEDKEY = 0x0001;
        private const UInt32 KEYEVENTF_KEYUP = 0x0002;

        [DllImport("user32.dll")]
        static extern void keybd_event(byte bVk, byte bScan, UInt32 dwFlags, UInt32 dwExtraInfo);

        [DllImport("user32.dll")]
        static extern Byte MapVirtualKey(UInt32 uCode, UInt32 uMapType);

        public static void VolumeUp()
        {
            keybd_event(VK_VOLUME_UP, MapVirtualKey(VK_VOLUME_UP, 0), KEYEVENTF_EXTENDEDKEY, 0);
            keybd_event(VK_VOLUME_UP, MapVirtualKey(VK_VOLUME_UP, 0), KEYEVENTF_EXTENDEDKEY | KEYEVENTF_KEYUP, 0);
        }

        public static void VolumeDown()
        {
            keybd_event(VK_VOLUME_DOWN, MapVirtualKey(VK_VOLUME_DOWN, 0), KEYEVENTF_EXTENDEDKEY, 0);
            keybd_event(VK_VOLUME_DOWN, MapVirtualKey(VK_VOLUME_DOWN, 0), KEYEVENTF_EXTENDEDKEY | KEYEVENTF_KEYUP, 0);
        }

        public static void Mute()
        {
            keybd_event(VK_VOLUME_MUTE, MapVirtualKey(VK_VOLUME_MUTE, 0), KEYEVENTF_EXTENDEDKEY, 0);
            keybd_event(VK_VOLUME_MUTE, MapVirtualKey(VK_VOLUME_MUTE, 0), KEYEVENTF_EXTENDEDKEY | KEYEVENTF_KEYUP, 0);
        }

        public static void NextSong()
        {
            keybd_event(VK_MEDIA_NEXT_TRACK, MapVirtualKey(VK_MEDIA_NEXT_TRACK, 0), KEYEVENTF_EXTENDEDKEY, 0);
            keybd_event(VK_MEDIA_NEXT_TRACK, MapVirtualKey(VK_MEDIA_NEXT_TRACK, 0), KEYEVENTF_EXTENDEDKEY | KEYEVENTF_KEYUP, 0);
        }

        public static void PreviousSong()
        {
            keybd_event(VK_MEDIA_PREV_TRACK, MapVirtualKey(VK_MEDIA_PREV_TRACK, 0), KEYEVENTF_EXTENDEDKEY, 0);
            keybd_event(VK_MEDIA_PREV_TRACK, MapVirtualKey(VK_MEDIA_PREV_TRACK, 0), KEYEVENTF_EXTENDEDKEY | KEYEVENTF_KEYUP, 0);
        }

        public static void PlayPauseSong()
        {
            keybd_event(VK_MEDIA_PLAY_PAUSE, MapVirtualKey(VK_MEDIA_PLAY_PAUSE, 0), KEYEVENTF_EXTENDEDKEY, 0);
            keybd_event(VK_MEDIA_PLAY_PAUSE, MapVirtualKey(VK_MEDIA_PLAY_PAUSE, 0), KEYEVENTF_EXTENDEDKEY | KEYEVENTF_KEYUP, 0);
        }

        public static void StopSong()
        {
            keybd_event(VK_MEDIA_STOP, MapVirtualKey(VK_MEDIA_STOP, 0), KEYEVENTF_EXTENDEDKEY, 0);
            keybd_event(VK_MEDIA_STOP, MapVirtualKey(VK_MEDIA_STOP, 0), KEYEVENTF_EXTENDEDKEY | KEYEVENTF_KEYUP, 0);
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

            if (!semantics.ContainsKey("stop") && 
                !semantics.ContainsKey("playPause") &&
                !semantics.ContainsKey("mute") &&
                !semantics.ContainsKey("control") &&
                !semantics.ContainsKey("volume"))
            {
                label1.Text = "No info provided!";
            }
            else if (semantics.ContainsKey("stop"))
            {
                label1.Text = rawText;
                StopSong();
                Update();
                synth.Speak(rawText);
            }
            else if (semantics.ContainsKey("playPause"))
            {
                label1.Text = rawText;
                PlayPauseSong();
                Update();
                synth.Speak(rawText);
            }
            else if (semantics.ContainsKey("mute"))
            {
                label1.Text = rawText;
                Mute();
                Update();
                synth.Speak(rawText);
            }
            else if (semantics.ContainsKey("volume"))
            {
                label1.Text = rawText;
                if (semantics.ContainsKey("volume").Equals(semantics["up"].Value))
                {
                    VolumeUp();
                } else
                {
                    VolumeDown();
                }
                Update();
                synth.Speak(rawText);
            }
            else if (semantics.ContainsKey("control"))
            {
                label1.Text = rawText;
                if (semantics.ContainsKey("control").Equals(semantics["next"].Value))
                {
                    NextSong();
                }
                else
                {
                    PreviousSong();
                }
                Update();
                synth.Speak(rawText);
            }
        }

        private Grammar CreateAudioControllerGrammar()
        {
            synth.Speak("Creando ahora la gramática");

            /*
             * Pausar la canción
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
             * Parar la cancion
             * frase()
             * 
             * Silenciar el ordenador
             * frase()
             */

            GrammarBuilder la = "la";
            GrammarBuilder el = "el";
            GrammarBuilder a = "a";
            GrammarBuilder cancion = "cancion";
            GrammarBuilder ordenador = "ordenador";
            GrammarBuilder volumen = "volumen";
            GrammarBuilder siguiente = "siguiente";
            GrammarBuilder anterior = "anterior";


            //play pause
            Choices playPauseChoice = new Choices();
            SemanticResultValue playPauseResultValue =
                    new SemanticResultValue("Reproducir", "play");
            GrammarBuilder playPauseValueBuilder = new GrammarBuilder(playPauseResultValue);
            playPauseChoice.Add(playPauseValueBuilder);

            playPauseResultValue =
                   new SemanticResultValue("Continuar", "play");
            playPauseValueBuilder = new GrammarBuilder(playPauseResultValue);
            playPauseChoice.Add(playPauseValueBuilder);

            playPauseResultValue =
                   new SemanticResultValue("Pausar", "pause");
            playPauseValueBuilder = new GrammarBuilder(playPauseResultValue);
            playPauseChoice.Add(playPauseValueBuilder);

            SemanticResultKey playPauseResultKey = new SemanticResultKey("playPause", playPauseChoice);
            GrammarBuilder playPause = new GrammarBuilder(playPauseResultKey);

            GrammarBuilder phrasePlayPause = playPause;
            phrasePlayPause.Append(la);
            phrasePlayPause.Append(cancion);

            //Volumen
            Choices volumeChoice = new Choices();
            SemanticResultValue volumeResultValue =
                    new SemanticResultValue("Subir", "up");
            GrammarBuilder volumeValueBuilder = new GrammarBuilder(volumeResultValue);
            volumeChoice.Add(volumeValueBuilder);

            volumeResultValue =
                   new SemanticResultValue("Aumentar", "up");
            volumeValueBuilder = new GrammarBuilder(volumeResultValue);
            volumeChoice.Add(volumeValueBuilder);

            volumeResultValue =
                   new SemanticResultValue("Bajar", "down");
            volumeValueBuilder = new GrammarBuilder(volumeResultValue);
            volumeChoice.Add(volumeValueBuilder);

            volumeResultValue =
                   new SemanticResultValue("Reducir", "down");
            volumeValueBuilder = new GrammarBuilder(volumeResultValue);
            volumeChoice.Add(volumeValueBuilder);

            SemanticResultKey volumeResultKey = new SemanticResultKey("volume", volumeChoice);
            GrammarBuilder volume = new GrammarBuilder(volumeResultKey);

            GrammarBuilder phraseVolume = volume;
            phraseVolume.Append(el);
            phraseVolume.Append(volumen);

            //control
            Choices controlChoice = new Choices();
            SemanticResultValue controlResultValue =
                    new SemanticResultValue("Avanzar", "next");
            GrammarBuilder controlValueBuilder = new GrammarBuilder(controlResultValue);
            controlChoice.Add(controlValueBuilder);

            controlResultValue =
                   new SemanticResultValue("Saltar", "next");
            controlValueBuilder = new GrammarBuilder(controlResultValue);
            controlChoice.Add(controlValueBuilder);

            controlResultValue =
                   new SemanticResultValue("Volver", "previous");
            controlValueBuilder = new GrammarBuilder(controlResultValue);
            controlChoice.Add(controlValueBuilder);

            controlResultValue =
                   new SemanticResultValue("Retroceder", "previous");
            controlValueBuilder = new GrammarBuilder(controlResultValue);
            controlChoice.Add(controlValueBuilder);

            SemanticResultKey controlResultKey = new SemanticResultKey("control", controlChoice);
            GrammarBuilder control = new GrammarBuilder(controlResultKey);

            GrammarBuilder phraseControl = control;
            phraseControl.Append(a);
            phraseControl.Append(la);
            Choices laElChoices = new Choices(siguiente, anterior);
            phraseControl.Append(laElChoices);
            phraseControl.Append(cancion);

            //mute
            Choices muteChoice = new Choices();
            SemanticResultValue muteResultValue =
                    new SemanticResultValue("Silenciar", "mute");
            GrammarBuilder muteValueBuilder = new GrammarBuilder(muteResultValue);
            muteChoice.Add(muteValueBuilder);

            SemanticResultKey muteResultKey = new SemanticResultKey("mute", muteChoice);
            GrammarBuilder mute = new GrammarBuilder(muteResultKey);

            GrammarBuilder phraseMute = mute;
            phraseMute.Append(el);
            phraseMute.Append(ordenador);

            //stop
            Choices stopChoice = new Choices();
            SemanticResultValue stopResultValue =
                    new SemanticResultValue("Parar", "stop");
            GrammarBuilder stopValueBuilder = new GrammarBuilder(stopResultValue);
            stopChoice.Add(stopValueBuilder);

            SemanticResultKey stopResultKey = new SemanticResultKey("stop", stopChoice);
            GrammarBuilder stop = new GrammarBuilder(stopResultKey);

            GrammarBuilder phraseStop = stop;
            phraseStop.Append(la);
            phraseStop.Append(cancion);

            Choices allPhrasesChoices = new Choices(new GrammarBuilder[] { phrasePlayPause, phraseVolume, phraseControl, phraseMute, phraseStop });

            Grammar grammar = new Grammar(allPhrasesChoices);
            grammar.Name = "Audio controller";
            return grammar;
        }

    }
}