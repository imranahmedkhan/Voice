using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Speech.Synthesis;
using System.Runtime.InteropServices;
using System.IO;

namespace voice
{
    class Program
    {
        // todo - just passthrough unless numlock/capslock/scrolllock detected?

        [DllImport("user32.dll", CharSet = CharSet.Auto, ExactSpelling = true, CallingConvention = CallingConvention.Winapi)]
        public static extern short GetKeyState(int keyCode);
        
        static void Help()
        {
            Console.WriteLine("voice.exe v0.5 by Eli Fulkerson 15 Mar 2019");
            Console.WriteLine("Just a thin command line wrapper around System.Speech.Synthesis.");
            Console.WriteLine("http://www.elifulkerson.com for updates");
            Console.WriteLine("");
            Console.WriteLine("Usage:  voice [-v VOLUME] [-r RATE] [-n NAME] [-m] [-f] [-l] [-d] [-?]");
            Console.WriteLine(" -v X   : Speak at volume X, where X is 0 to 100");
            Console.WriteLine(" -r X   : Speak at rate X, where X is -10 to 10");
            Console.WriteLine(" -n \"X\" : Speak using named voice X");
            Console.WriteLine(" -l     : List availabled named voices");
            Console.WriteLine(" -m     : Just use a male voice, if available");
            Console.WriteLine(" -f     : Just use a female voice, if available");
            Console.WriteLine(" -d     : (Optional) Indicate that arguments are done.");
            Console.WriteLine("          ... All text to the right of -d will be spoken even if it contains valid arguments.");
            Console.WriteLine(" -i     : Interactive mode");
            Console.WriteLine(" -t     : In pipe mode, output text to console as well as read it aloud");
            Console.WriteLine(" -s     : In pipe mode, copy to stderr and read aloud only when scroll lock is enabled");
            Console.WriteLine(" -p     : Output progress information to stdout (probably don't use with -t, -s)");
            Console.WriteLine("         (Format is - CharPos:CharCount:AudioPos:Text)");
            Console.WriteLine(" -k X   : Read from input filename X");
            Console.WriteLine(" -o X   : Output to filename X in .WAV format (overwriting any previous file)");
            Console.WriteLine(" --mono : By default, we output stereo (2 channels).  This flag outputs mono (1 channel)");
            Console.WriteLine(" --8bit : By default, we output 16 bit samples.  This flag outputs 8 bit.)");
            Console.WriteLine(" --khz X: Output at X khz.  Default 44.");
            Console.WriteLine(" -?     : Help and version information");
            Console.WriteLine("");
            Console.WriteLine("Optional:");
            Console.WriteLine("For the sake of usability and/or self documentation, the *actual* arguments accepted are...");
            Console.WriteLine(" -v or /v or /volume or --volume");
            Console.WriteLine(" -r or /r or /rate or --rate");
            Console.WriteLine(" -n or /n or /name or --name");
            Console.WriteLine(" -l or /l or /list or --list");
            Console.WriteLine(" -m or /m or /male or --male");
            Console.WriteLine(" -f or /f or /female or --female");
            Console.WriteLine(" -d or /d or /done or --done");
            Console.WriteLine(" -i or /i or /interactive or --interactive");
            Console.WriteLine(" -t or /t or /tee or --tee");
            Console.WriteLine(" -s or /s or /scroll or --scroll");
            Console.WriteLine(" -p or /p or /progress or --progress");
            Console.WriteLine(" -k or /k or /input or --input");
            Console.WriteLine(" -o or /o or /output or --output");
            Console.WriteLine("  --mono or /mono");
            Console.WriteLine("  --8bit or /8bit");
            Console.WriteLine("  --khz or /khz");
            Console.WriteLine(" -? or --? or /? or /h or -h or /help or --help or -help or nothing at all");
            
        }

        // Write each word and its character postion to the console.
        static void report_progress(object sender, SpeakProgressEventArgs e)
        {
            Console.WriteLine("{0}:{1}:{2}:\"{3}\"",e.CharacterPosition, e.CharacterCount, e.AudioPosition, e.Text);
        }

        static void Main(string[] args)
        {

            SpeechSynthesizer s = new SpeechSynthesizer();

            int Volume = s.Volume;
            int Rate = s.Rate;
            string line;
            bool mode_interactive = false;
            bool mode_tee = false;
            bool mode_scroll_inspect = false;
            bool mode_report_progress = false;

            int consumed = 0;  // the number of args that are used up - and thusly not the payload

            string InputFilename = "";
            string OutputFilename = "";

            // sound properties
            System.Speech.AudioFormat.AudioChannel numchannels = System.Speech.AudioFormat.AudioChannel.Stereo;
            System.Speech.AudioFormat.AudioBitsPerSample numbits = System.Speech.AudioFormat.AudioBitsPerSample.Sixteen;
            Int32 khz = 44;

            // So... unicode.  Notepad for instance can save accented characters in ANSI mode but you really want utf-8
            bool unicode_warning_sent = false;

            // This annoys me that I had to do this.
            bool keyAvailableHack = false;
            try
            {
                keyAvailableHack = System.Console.KeyAvailable;
            }
            catch
            {
                keyAvailableHack = true;
            }


            // help for no args unless we are piping
            if (args.Length == 0 && keyAvailableHack == false)
            {
                Help();
                Environment.Exit(0);
            }

            for (int i = 0; i < args.Length; i++)
            {
                if (args[i] == "-v" || args[i] == "--volume" || args[i] == "/v" || args[i] == "/volume")
                {
                    try
                    {
                        Volume = Convert.ToInt32(args[i + 1]);
                        if (Volume < 0 || Volume > 100)
                        {
                            Console.WriteLine("Volume must be between 0 and 100, inclusive.");
                            Environment.Exit(1);
                        }
                    }
                    catch
                    {
                        Console.WriteLine("Invalid volume specified");
                        Environment.Exit(1);
                    }

                    i++;
                    consumed += 2;
                    
                }

                if (args[i] == "-r" || args[i] == "--rate" || args[i] == "/r" || args[i] == "/rate")
                {
                    try
                    {
                        Rate = Convert.ToInt32(args[i + 1]);
                        if (Rate < -10 || Rate > 10)
                        {
                            Console.WriteLine("Rate must be between -10 and 10, inclusive.");
                            Environment.Exit(1);
                        }
                    }
                    catch
                    {
                        Console.WriteLine("Invalid rate specified");
                        Environment.Exit(1);
                    }

                    i++;
                    consumed += 2;

                }

                if (args[i] == "-m" || args[i] == "--male" || args[i] == "/m" || args[i] == "/male")
                {
                    s.SelectVoiceByHints(VoiceGender.Male);
                    consumed++;
                }

                if (args[i] == "-f" || args[i] == "--female" || args[i] == "/f" || args[i] == "/female")
                {
                    s.SelectVoiceByHints(VoiceGender.Female);
                    consumed++;
                }

                if (args[i] == "-n" || args[i] == "--name" || args[i] == "/n" || args[i] == "/name")
                {
                    try
                    {
                        s.SelectVoice(args[i + 1]);
                    }
                    catch
                    {
                        Console.WriteLine("Invalid voice name specified.  Did you use quotes?  Try using --list to see available names");
                        Environment.Exit(1);
                    }
                    i++;
                    consumed+=2;
                }

                if (args[i] == "-l" || args[i] == "--list" || args[i] == "/l" || args[i] == "/list")
                {
                    //choose by name
                    foreach (InstalledVoice v in s.GetInstalledVoices())
                    {
                        Console.WriteLine("\"{0}\" - {1},{2},{3}", v.VoiceInfo.Name, v.VoiceInfo.Age, v.VoiceInfo.Gender, v.VoiceInfo.Culture);
                        //Console.WriteLine(v.VoiceInfo.Description);
                        //Console.WriteLine(v.VoiceInfo.Id);
                    }
                    Environment.Exit(0);
                }

                if (args[i] == "-?" || args[i] == "--?" || args[i] == "/?" || args[i] == "/h" || args[i] == "-h" || args[i] == "/help" || args[i] == "--help" || args[i] == "-help")
                {
                    //help and version and quit
                    Help();
                    Environment.Exit(0);
                }

                if (args[i] == "-i" || args[i] == "--interactive" || args[i] == "/i" || args[i] == "/interactive")
                {
                    mode_interactive = true;
                    consumed++;
                }

                if (args[i] == "-t" || args[i] == "--tee" || args[i] == "/t" || args[i] == "/tee")
                {
                    mode_tee = true;
                    consumed++;
                }

                if (args[i] == "-s" || args[i] == "--scroll" || args[i] == "/s" || args[i] == "/scroll")
                {
                    mode_scroll_inspect = true;
                    consumed++;
                }

                if (args[i] == "-p" || args[i] == "--progress" || args[i] == "/p" || args[i] == "/progress")
                {
                    mode_report_progress = true;
                    consumed++;
                }

                if (args[i] == "-d" || args[i] == "--done" || args[i] == "/d" || args[i] == "/done")
                {
                    // nothing after "done" gets parsed, in case there are actual codes in the speech
                    consumed++;
                    break;
                }

                if (args[i] == "-k" || args[i] == "--input" || args[i] == "/k" || args[i] == "/input")
                {
                    InputFilename = args[i + 1];
                    i++;
                    consumed += 2;
                }

                if (args[i] == "-o" || args[i] == "--output" || args[i] == "/o" || args[i] == "/output")
                {
                    OutputFilename = args[i + 1];
                    i++;
                    consumed += 2;
                }


                if (args[i] == "--mono" || args[i] == "/mono")
                {
                    numchannels = System.Speech.AudioFormat.AudioChannel.Mono;
                    consumed++;
                    
                }

                if (args[i] == "--8bit" || args[i] == "/8bit")
                {
                    numbits = System.Speech.AudioFormat.AudioBitsPerSample.Eight;
                    consumed++;
                    
                }

                if (args[i] == "--khz" || args[i] == "/khz")
                {
                    khz = Convert.ToInt32(args[i + 1]);
                    i++;
                    consumed+=2;   
                }
            }

            if (mode_report_progress)
            {
                s.SpeakProgress += new EventHandler<SpeakProgressEventArgs>(report_progress);
            }

            s.Volume = Volume;
            s.Rate = Rate;

            System.Speech.AudioFormat.SpeechAudioFormatInfo fmt = new System.Speech.AudioFormat.SpeechAudioFormatInfo(khz*1000, numbits, numchannels);
            
            if (OutputFilename != "")
            {
                try
                {
                    s.SetOutputToWaveFile(OutputFilename, fmt);   
                }
                catch (Exception e)
                {
                    Console.WriteLine("Error, the output file could not be written.");
                    Console.WriteLine(e.Message);
                }
            }
            
            if (InputFilename != "")
            {
                try
                {
                    using (StreamReader sr = new StreamReader(InputFilename, true))
                    {
                        while ((line = sr.ReadLine()) != null) {

                            // Including this warning because it is apparently very easy (I did it, anyway) to save what you *think* are properly
                            // accented non-english characters in a notepad file under ANSI - they show up properly in notepad etc but when you 
                            // feed them to Speech.Synthesis they are replaced with the 65533 garbage character.
                            if (unicode_warning_sent == false) {
                                for (int i = 0; i < line.Length; i++) {
                                    if (Convert.ToInt32(line[i]) == 65533)
                                    {
                                        Console.WriteLine("Warning:  Unicode replacement character 65533 detected.  Check encoding on input file.");
                                        unicode_warning_sent = true;
                                        break;
                                    }
                                }
                            }
                            s.Speak(line);      
                        }
                    }
                }
                catch (Exception e)
                {
                    Console.WriteLine("The input file could not be read:");
                    Console.WriteLine(e.Message);
                }
            }
            
            // "from stdin, noninteractive" mode
            if (keyAvailableHack)
            {
                while (true)
                {
                    line = Console.ReadLine();
                    if (line == null)
                    {
                        break;
                    }
                    else
                    {
                        if (mode_scroll_inspect)
                        {
                            // check for scroll lock
                            if ((((ushort)GetKeyState(0x91)) & 0xffff) != 0)
                            {
                                Console.Error.WriteLine(line);
                                s.Speak(line);
                            }
                            Console.WriteLine(line);
                            continue;
                        } 

                        if (mode_tee)
                        {
                            Console.WriteLine(line);
                        }
                        s.Speak(line);
                    }
                }
                Environment.Exit(0);
            }
            
            // from stdin, interactive
            // "from stdin, noninteractive" mode


            if (mode_interactive)
            {
                Console.WriteLine("Interactive mode.  Control-c to quit.");
                while (true)
                {
                    line = Console.ReadLine();
                    if (line != null)
                    {

                        s.Speak(line);

                    }
                    else
                    {
                        break;
                    }
                }
                Environment.Exit(0);
            }

            // "from the args" mode
            string tosay = String.Join(" ", args.Skip(consumed));
            if (tosay.Length > 0)
            {
                s.Speak(tosay);

            }


            Environment.Exit(0);
         
        }
    }
}
