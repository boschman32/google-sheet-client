using System;
using System.Text;
using System.Threading;

namespace GoogleSheetClient
{
    //From: https://gist.github.com/DanielSWolf/0ab6a96899cc5377bf54
    //MIT LICENSE
    public class ProgressBar : IDisposable, IProgress<double>
    {
        private const int BAR_LENGTH = 50;
        private readonly TimeSpan AnimationInterval = TimeSpan.FromSeconds(1.0 / 8);
        private const string Animation = @"|/-\";
        private int AnimationIndex = 0;

        private readonly Timer Timer;

        private double CurrentProgress = 0;
        private string CurrentText = string.Empty;
        private bool Disposed = false;

        private string Prefix;

        public ProgressBar(string prefix = "")
        {
            Prefix = prefix;

            Timer = new Timer(TimerHandler);

            // A progress bar is only for temporary display in a console window.
            // If the console output is redirected to a file, draw nothing.
            // Otherwise, we'll end up with a lot of garbage in the target file.
            if (!Console.IsOutputRedirected)
            {
                ResetTimer();
            }
        }

        public void Report(double value)
        {
            value = Math.Max(0, Math.Min(1, value));
            Interlocked.Exchange(ref CurrentProgress, value);
        }

        private void TimerHandler(object state)
        {
            lock (Timer)
            {
                if (Disposed)
                {
                    return;
                }

                int progressCount = (int)(CurrentProgress * BAR_LENGTH);
                int percent = (int)(CurrentProgress * 100);
                string text = string.Format("{0}[{1}{2}] {3,4}% {4}"
                    , Prefix
                    , new string('#', progressCount)
                    , new string('-', BAR_LENGTH - progressCount)
                    , percent
                    , Animation[AnimationIndex++ % Animation.Length]);
                UpdateBar(text);

                ResetTimer();
            }
        }

        private void UpdateBar(string text)
        {
            // Get length of common portion
            int commonPrefixLength = 0;
            int commonLength = Math.Min(CurrentText.Length, text.Length);
            while (commonPrefixLength < commonLength && text[commonPrefixLength] == CurrentText[commonPrefixLength])
            {
                commonPrefixLength++;
            }

            // Backtrack to the first differing character
            StringBuilder outputBuilder = new StringBuilder();
            outputBuilder.Append('\b', CurrentText.Length - commonPrefixLength);

            // Output new suffix
            outputBuilder.Append(text.Substring(commonPrefixLength));

            // If the new text is shorter than the old one: delete overlapping characters
            int overlapCount = CurrentText.Length - text.Length;
            if (overlapCount > 0)
            {
                outputBuilder.Append(' ', overlapCount);
                outputBuilder.Append('\b', overlapCount);
            }

            Console.Write(outputBuilder);
            CurrentText = text;
        }

        private void ResetTimer()
        {
            Timer.Change(AnimationInterval, TimeSpan.FromMilliseconds(-1));
        }

        public void Dispose()
        {
            lock (Timer)
            {
                Disposed = true;
                UpdateBar(string.Empty);
            }
        }
    }

}
