using System;
using System.Collections.Generic;
using System.Diagnostics;
using CSCore.CoreAudioAPI;

namespace SpotifyVolumeManagerPrototype
{
    class Program
    {
        const float SpotifyMax           = 0.5F;
        const float SpotifyMin           = 0.05F;
        const float VolumeIncrements     = 0.05F;

        static float threshold           = 0.02F;

        static void Main(string[] args)
        {
            // Required to see if the processes are even open
            int spotify_pID = 0;
            int teams_pID = 0;

            AudioSessionControl spotify_AudioStream = null;
            // Teams needs to be a List because there are multiple Teams audio streams at one time
            List<AudioSessionControl> teams_AudioStream = new List<AudioSessionControl>();


            // Change threshold if one has been given
            if (args.Length == 1)
            {
                if (!float.TryParse(args[0], out threshold))
                {
                    Console.WriteLine("ERR: Please enter a decimal argument.");
                    Console.WriteLine("Usage: a.exe <threshold value>");
                    return;
                } else if (threshold < 0F | threshold > 1F)
                {
                    Console.WriteLine("ERR: Please enter a decimal argument larger than 0 and smaller than 1.");
                    Console.WriteLine("Usage: a.exe <threshold value between 0.0 and 1.0>");
                    return;
                }
            } else
            {
                Console.WriteLine("WARN: No threshold provided, running with default threshold {0}", threshold);
            }

            // Get Process ID of teams and spotify
            foreach (var process in Process.GetProcesses())
            {
                if (process.ProcessName == "Spotify" && !string.IsNullOrEmpty(process.MainWindowTitle))
                {
                    spotify_pID = process.Id;
                } else if (process.ProcessName == "Teams" && !string.IsNullOrEmpty(process.MainWindowTitle))
                {
                    teams_pID = process.Id;
                }
            }

            // Quit application if one of the applications were not found
            if (spotify_pID == 0 | teams_pID == 0)
            {
                Console.WriteLine("ERROR: Spotify or Teams could not be found! Please open the missing application.");
                Console.ReadKey();
                return;
            }

            // Get AudioSessionControls of Teams and Spotify by going through 
            using (var sessionManager = GetDefaultAudioSessionManager2(DataFlow.Render))
            {
                using (var sessionEnumerator = sessionManager.GetSessionEnumerator())
                {
                    foreach (var session in sessionEnumerator)
                    {
                        string name;

                        // Get the name of the process that is creating the audio stream
                        using (var audioSessionControl2 = session.QueryInterface<AudioSessionControl2>())
                        {
                            var process = audioSessionControl2.Process;
                            name = audioSessionControl2.DisplayName;
                            
                            if (process != null)
                            {
                                if (name == "") { name = process.ProcessName; }
                            }

                            if (name == "") { name = "Unnamed"; }
                        }

                        if (name == "Spotify")
                        {
                            spotify_AudioStream = session;
                        } else if (name == "Teams")
                        {
                            teams_AudioStream.Add(session);
                        }
                    }
                }
            }

            // One of them wasn't found, send error
            if (spotify_AudioStream == null | teams_AudioStream.Count == 0)
            {
                Console.WriteLine("ERROR: Teams or Spotify is not outputting an audio stream, please start a call on Teams or play music on Spotify!");
                Console.WriteLine("\nPress any key to continue...");
                Console.ReadKey();
                return;
            }

            // Show audio levels every 200ms
            while (true)
            {  
                // spotify
                if (spotify_AudioStream.QueryInterface<AudioMeterInformation>().PeakValue < Math.Pow(10, -4))
                {
                    Console.WriteLine("Spotify is not outputting anything.");
                } /* else
                {
                    Console.WriteLine("Spotify Audio Level: {0}", spotify_AudioStream.QueryInterface<AudioMeterInformation>().PeakValue);
                } */

                float teams_vol = 0F;
                // Ignore first audiostream of teams bc it's for notifications, not actual audio during a meeting
                for (int i = 1; i < teams_AudioStream.Count; i++)
                {
                    teams_vol += teams_AudioStream[i].QueryInterface<AudioMeterInformation>().PeakValue;
                }

                if (teams_vol < Math.Pow(10, -5))
                {
                    Console.WriteLine("Teams is not outputting anything.");
                } else
                {
                    teams_vol /= teams_AudioStream.Count;
                    // Console.WriteLine("Teams Audio Level: {0}", teams_vol);
                } 

                
                if (teams_vol > threshold)
                {
                    if (spotify_AudioStream.QueryInterface<SimpleAudioVolume>().MasterVolume != SpotifyMin)
                    {
                        // Someone is likely speaking on teams and we did not increase before, lower volume
                        Console.WriteLine("Lowering spotify volume, Teams volume: {0}...", teams_vol);
                        ChangeVolumeSlowlyCSCore(spotify_AudioStream.QueryInterface<SimpleAudioVolume>(), SpotifyMin);
                        Console.WriteLine("Done lowering spotify volume...");

                        // Delay to not have it go up and down weirdly
                        System.Threading.Thread.Sleep(3000);
                    }
                } else
                {
                    if (spotify_AudioStream.QueryInterface<SimpleAudioVolume>().MasterVolume != SpotifyMax)
                    {
                        // Nobody is speaking on teams and we did not decease before, increase volume
                        Console.WriteLine("Increasing spotify volume, Teams volume: {0}...", teams_vol);
                        ChangeVolumeSlowlyCSCore(spotify_AudioStream.QueryInterface<SimpleAudioVolume>(), SpotifyMax);
                        Console.WriteLine("Done increasing spotify volume...");

                        // Delay to not have it go up and down weirdly
                        System.Threading.Thread.Sleep(500);
                    }
                }

                System.Threading.Thread.Sleep(200);
            }
        }
       
        private static void ChangeVolumeSlowlyCSCore(SimpleAudioVolume simpleAudioVolume, float vol)
        {
            float current_vol = simpleAudioVolume.MasterVolume;
            
            while ((float) Math.Round(simpleAudioVolume.MasterVolume, 2) != vol)
            {
                current_vol = simpleAudioVolume.MasterVolume;

                if (current_vol > vol)
                {
                    // lower volume
                    if (current_vol - vol < VolumeIncrements)
                    {
                        // Just change it to the volume because using the increments will cause a loop
                        simpleAudioVolume.MasterVolume = vol;
                    } else
                    {
                        // Increment volume
                        simpleAudioVolume.MasterVolume = (float)current_vol - VolumeIncrements;
                    }
                }
                else if (current_vol < vol)
                {
                    // increase volume
                    if (vol - current_vol < VolumeIncrements)
                    {
                        // Just change it to the volume because using the increments will cause a loop
                        simpleAudioVolume.MasterVolume = vol;
                    }
                    else
                    {
                        // Increment volume
                        simpleAudioVolume.MasterVolume = (float)current_vol + VolumeIncrements;
                    }
                }

                // Delay to give illusion of smoothness
                System.Threading.Thread.Sleep(20);
            }
        }
        
        private static AudioSessionManager2 GetDefaultAudioSessionManager2(DataFlow dataFlow)
        {
            using (var enumerator = new MMDeviceEnumerator())
            {
                using (var device = enumerator.GetDefaultAudioEndpoint(dataFlow, Role.Multimedia))
                {
                    return AudioSessionManager2.FromMMDevice(device);
                }
            }
        }
    }
}




