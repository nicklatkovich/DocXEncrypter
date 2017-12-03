using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace DocxEncrypter {
    public static class Utils {

        static Random Random = new Random( );
        
        public static bool[ ] ConvertToBits(this byte[ ] bytes) {
            var result = new bool[bytes.Length * 8];
            for (var byteIndex = 0; byteIndex < bytes.Length; byteIndex++) {
                var @byte = bytes[byteIndex];
                for (ushort bitIndex = 0; bitIndex < 8; bitIndex++) {
                    result[byteIndex * 8 + bitIndex] = (@byte & (1 << bitIndex)) == 0 ? false : true;
                }
            }
            return result;
        }

        public static int[ ] GetShuffled(int size) {
            var result = new int[size];
            for (var i = 0; i < size; i++) {
                result[i] = i;
            }
            return result.Shuffle( );
        }

        public static T[ ] Swap<T>(this T[ ] target, int index1, int index2) {
            T swap = target[index1];
            target[index1] = target[index2];
            target[index2] = swap;
            return target;
        }

        public static T[ ] Shuffle<T>(this T[ ] target, Random random = null) {
            if (random == null) {
                random = Random;
            }
            for (var i = 0; i < target.Length; i++) {
                var j = random.Next(target.Length);
                target.Swap(i, j);
            }
            return target;
        }

    }
}
