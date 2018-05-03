using System;

namespace Microsoft.Operations
{
    /// <summary>
    /// Maths and other number functions
    /// </summary>
    public static class Maths
    {
        private static readonly Random random = new Random();
        private static readonly object syncLock = new object();

        /// <summary>
        /// Provides a random number from the internal generator. Preferred method because it
        /// provides TRUE random numbers.
        /// </summary>
        public static int RandomNumber(int min, int max)
        {
            lock (syncLock)
            {
                return random.Next(min, max);
            }
        }
    }
}