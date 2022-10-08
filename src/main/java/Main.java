import org.apache.poi.poifs.crypt.Decryptor;
import org.apache.poi.poifs.crypt.EncryptionInfo;
import org.apache.poi.poifs.filesystem.POIFSFileSystem;

import java.io.File;
import java.io.IOException;
import java.security.GeneralSecurityException;
import java.time.Duration;
import java.time.Instant;
import java.util.ArrayList;
import java.util.Arrays;
import java.util.concurrent.*;
import java.util.stream.IntStream;

public class Main {
    public static void main(String[] args) {
        String fileToCrack = (args.length > 0 && args[1] != null) ? args[1] : "src/test/resources/test_file.xlsx";
        final File inputFile = new File(fileToCrack);
        String crackedPassword = crackPassword(inputFile);
        System.out.println("Password found: " + crackedPassword);
    }

    public static String crackPassword(File inputFile) {
        char[] charSet = prepareCrackingCharacterSet();

        int minPasswordLength = 2;
        int maxPasswordLength = 16;

        int threadCount = Runtime.getRuntime().availableProcessors();

        ExecutorService threadPoolExecutor = new ThreadPoolExecutor(threadCount, threadCount, 0L, TimeUnit.MILLISECONDS, new LinkedBlockingQueue<>());
        CompletableFuture<String> cf = new CompletableFuture<>();
        BlockingQueue<String> passwordQueue = new LinkedBlockingQueue<>(threadCount * 32);

        final Decryptor excelDecryptor = getDecryptor(inputFile);

        Runnable producer = passwordProvider(charSet, passwordQueue, minPasswordLength, maxPasswordLength);
        Runnable consumer = passwordCracker(cf, passwordQueue, excelDecryptor);

        executeRunnableDesiredTimes(1, threadPoolExecutor, producer);
        executeRunnableDesiredTimes(threadCount - 3, threadPoolExecutor, consumer);

        return doCrackPassword(threadPoolExecutor, cf);
    }

    private static char[] prepareCrackingCharacterSet() {
        ArrayList<Character> characters = new ArrayList<>();

        IntStream.range(97, 123).forEach(i -> characters.add((char)i));
        //IntStream.range(97, 127).forEach(i -> characters.add((char)i));
        //IntStream.range(65, 97).forEach(i -> characters.add((char)i));
        //IntStream.range(33, 65).forEach(i -> characters.add((char)i));

        return getCharSet(characters);
    }

    private static String doCrackPassword(ExecutorService service, CompletableFuture<String> cf) {
        String result = "";

        try {
            Instant start = Instant.now();
            result = cf.get();
            Instant finish = Instant.now();
            long timeElapsed = Duration.between(start, finish).toMillis();
            System.out.println("Total time elapsed: " + timeElapsed + "ms");
            service.shutdownNow();
        } catch (InterruptedException | ExecutionException e) {
            e.printStackTrace();
        }

        return result;
    }

    private static void executeRunnableDesiredTimes(int times, ExecutorService service, Runnable runnable) {
        for (int i = 0; i < times; ++i) {
            service.execute(runnable);
        }
    }

    private static Runnable passwordProvider(char[] charSet, BlockingQueue<String> passwordQueue, int minLen, int maxLen) {
        return () -> {
            char[] result = new char[maxLen];
            int[] index = new int[maxLen];

            int charSetSize = charSet.length;

            Arrays.fill(result, 0, maxLen, charSet[0]);
            Arrays.fill(index, 0, maxLen, 0);

            for (int i = minLen; i <= maxLen; i++) {
                int updateIndex = 0;

                do {
                    try {
                        passwordQueue.offer(new String(result, 0, i), 12, TimeUnit.HOURS);
                    } catch (InterruptedException e) {
                        break;
                    }

                    for (updateIndex = i-1;
                         updateIndex != -1 && ++index[updateIndex] == charSetSize;
                         result[updateIndex] = charSet[0], index[updateIndex] = 0, updateIndex--);

                    if(updateIndex != -1) {
                        result[updateIndex] = charSet[index[updateIndex]];
                    }
                } while (updateIndex != -1);
            }
        };
    }

    private static Runnable passwordCracker(CompletableFuture<String> cf, BlockingQueue<String> passwordQueue, Decryptor excelDecryptor) {
        return () -> {
            while (!Thread.interrupted()) {
                try {
                    String password = passwordQueue.take();
                    //System.out.println("Testing password: " + password);
                    boolean decryptResult = excelDecryptor.verifyPassword(password);
                    if (decryptResult) {
                        cf.complete(password);
                        break;
                    }
                } catch (InterruptedException | GeneralSecurityException ignored) {
                    break;
                }
            }
        };
    }

    private static Decryptor getDecryptor(File inputFile) {
        POIFSFileSystem fileSystem;
        EncryptionInfo info;
        Decryptor decryptor = null;

        try {
            fileSystem = new POIFSFileSystem(inputFile);
            info = new EncryptionInfo(fileSystem);
            decryptor = Decryptor.getInstance(info);
        } catch (IOException e) {
            e.printStackTrace();
        }

        return decryptor;
    }

    private static char[] getCharSet(ArrayList<Character> letters) {
        char[] chars = new char[letters.size()];
        for (int i = 0; i < chars.length; i++) {
            chars[i] = letters.get(i);
        }
        return chars;
    }
}
