import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.util.CellReference;
import java.util.*;
import javax.net.ssl.HttpsURLConnection;
import java.io.*;
import java.net.*;


public class Main {
    public static final String RESET = "\u001B[0m";
    public static final String VERMELHO = "\u001B[31m";
    public static final String VERDE = "\u001B[32m";
    public static final String AMARELO = "\u001B[33m";
    public static final String CIANO = "\u001B[36m";

    public static void main(String[] args) {
        String versaoLocal = "1.1";
        System.out.println("🔄 A verificar existençia de atualizações na nuvem ");
        String versaoNuvem = ObterVersaoNuvem();
        
        if(versaoNuvem !=null){
            if(!versaoLocal.equals(versaoNuvem)){
                System.out.println("📢 Nova versão existente: (Nuvem: " + versaoNuvem + " | Local: " + versaoLocal + ")");
                System.out.println("⏳ A iniciar download da nova versão...");

                String urlBase = "https://raw.githubusercontent.com/anonimoani31-rgb/Comparador-Updates/refs/heads/main/";

                descarregarAtualizacao(urlBase + "Main.class", "Main.class");
                descarregarAtualizacao(urlBase + "Main$1.class", "Main$1.class");

                System.out.println("programa atualizado com exito!! Por favor reinicie o programa para aplicar as alterações.");
            }else{
                System.out.println("✅ O programa foi atualizado para a versão: " + versaoLocal);
            }
        }
        Scanner LeituraTeclado = new Scanner(System.in);

        System.out.println(AMARELO + "=== SISTEMA DE CONFIGURAÇÃO DE CAMINHOS ===" + RESET);

        System.out.println("Insira o caminho completo do primeiro ficheiro ");
        String PrimeiroCamimnho = LeituraTeclado.nextLine().trim();
        System.out.println("insira o caminho do segundo ficheiro");
        String SegundoCaminho = LeituraTeclado.nextLine().trim();

        File f1 = new File(PrimeiroCamimnho);
        File f2 = new File(SegundoCaminho);

        if(!f1.exists() || !f2.exists()){
            System.out.println(VERMELHO + "ERRO CRITÍCO: Um ou ambos os ficheiros Excel não foram encontrados no caminho especificado!" + RESET);
            System.out.println("Por favor, verifique se os nomes e caminhos estão corretos.");
            return;
        }

        if(f1.isDirectory() || f2.isDirectory()){
            System.out.println(VERMELHO + "ERRO CRITÍCO: O caminho fornecido aponta para uma pasta e não para um ficheiro .xlsx!" +  RESET);
        } else {
            System.out.println(VERDE + "✓ Ficheiros encontrados e validados com sucesso!✓ A iniciar processamento..." + RESET);
        }

        boolean ignorarMaiusculas = false;

        System.out.println("Deseja ignorar miusculas e minusculas na comparação (ex: tratar `texto` e `Texto` como iguais) [Sim ou Não]");
        String respostaCase = LeituraTeclado.nextLine().trim();

        if(respostaCase.equalsIgnoreCase("Sim") || respostaCase.equalsIgnoreCase("s")){
            ignorarMaiusculas = true;
            System.out.println(AMARELO + "Configuração Ativada : A Ingnorar maisculas/minusculas." + RESET);
        } else {
            System.out.println(AMARELO + "Comfiguração Ativada : A Diferençiar maiusculas/minusculas. (Case Sensitive)" + RESET);
        }

        System.out.println("A comparar ficheiros completos");
        String ExtençãoF1 = PrimeiroCamimnho.substring(PrimeiroCamimnho.lastIndexOf(".")).toLowerCase();
        String ExtençãoF2 = SegundoCaminho.substring(SegundoCaminho.lastIndexOf(".")).toLowerCase();

        if(!ExtençãoF1.equals(ExtençãoF2)){
            System.out.println(VERMELHO + "ERRO: Não é possivel comparar ficheiros de formatos diferentes (" + ExtençãoF1 + "vs" + ExtençãoF2 + ")" + RESET);
            return;
        }

        int numColunasF1 = 0;
        int numColunasF2 = 0;

        if(ExtençãoF1.equals(".xlsx") || ExtençãoF1.equals(".xls")){
            numColunasF1 = ObterTotalColunasExcel(PrimeiroCamimnho);
            numColunasF2 = ObterTotalColunasExcel(SegundoCaminho);
        } else if(ExtençãoF1.equals(".csv")){
            String sep = DescobrirSeparador(PrimeiroCamimnho);
            numColunasF1 = ObterTotalColunasCSV(PrimeiroCamimnho, sep);
            numColunasF2 = ObterTotalColunasCSV(SegundoCaminho, sep);
        }

        if (numColunasF1 != numColunasF2) {
            System.out.println(VERMELHO + "⚠️ ATENÇÃO EMPRESARIAL: Os ficheiros têm estruturas de colunas diferentes!" + RESET);
            System.out.printf("   Ficheiro 1: %d colunas | Ficheiro 2: %d colunas.%n", numColunasF1, numColunasF2);
            System.out.println("Deseja forçar a comparação mesmo assim? [Sim/Não]");
            String continuar = LeituraTeclado.nextLine().trim();
            
            if (!continuar.equalsIgnoreCase("sim") && !continuar.equalsIgnoreCase("s")) {
                System.out.println("Operação cancelada pelo utilizador de forma segura.");
                return;
            }
        }
            
        switch (ExtençãoF1){
            case ".xlsx":
            case "xls":
                System.out.println("Detetado formato Excel. A iniciar motor Apache POI...");
                fazerComparação(PrimeiroCamimnho, SegundoCaminho, new ArrayList<>(), ignorarMaiusculas);
                break;

            case ".csv":
                System.out.println("detectado formato CSV. A iniciar motor ultra-leve de texto");
                fazerComparaçãoCSV(PrimeiroCamimnho, SegundoCaminho, new ArrayList<>(), ignorarMaiusculas);
                break;

            default:
                System.out.println(VERMELHO + "O formato " + ExtençãoF1 + "náo é suportado por esta plataforma" + RESET);
        }

        mostrarMenuColunas(PrimeiroCamimnho);

        List<Integer> ColunasAignorar = new ArrayList<>();

        System.out.println("Insira as colunas que deseja ignorar na proxima comparação");
        String resposta = LeituraTeclado.nextLine();

        if(resposta.equalsIgnoreCase("sim") || resposta.equalsIgnoreCase("s")){
            System.out.println("Insira o numero das colunas que deseja ignorar separadas por espaço (ex: 0 2):");
            System.out.println("(Lembrete: Coluna A=0, B = 1, C = 2...)");
            
            String[] numeros = LeituraTeclado.nextLine().split(" ");
            for (String num : numeros){
                ColunasAignorar.add(Integer.parseInt(num));
            }

            System.out.println("A realizar Re-comparação(A IGNORAR COLUNAS ESCOLHIDAS)");
            fazerComparação(PrimeiroCamimnho, SegundoCaminho, ColunasAignorar, ignorarMaiusculas);
        } else {
            System.out.println("Re-comparação executada com exito - programa terminado");
        }
    } // Aqui fecha corretamente o método main!

    public static String ObterVersaoNuvem(){
        try{
            String urlString = "https://raw.githubusercontent.com/anonimoani31-rgb/Comparador-Updates/refs/heads/main/versao.txt";

            URL url = new URL(urlString);
            HttpURLConnection conexao = (HttpURLConnection) url.openConnection();
            conexao.setRequestMethod("GET");
            
            BufferedReader leitor = new BufferedReader(new InputStreamReader(conexao.getInputStream()));
            String versao = leitor.readLine();
            leitor.close();

            return versao.trim();
        } catch (Exception e){
            System.out.println("⚠️ Não foi possível verificar atualizações (Sem internet)."); 
            return null;   
        }
    }

    public static void descarregarAtualizacao(String ulrlFicheiro, String nomeDestino){
        try{
            URL url = new URL(ulrlFicheiro);
            HttpsURLConnection conexao = (HttpsURLConnection) url.openConnection();

            InputStream entrada = conexao.getInputStream();
            FileOutputStream saida = new FileOutputStream(nomeDestino);

            byte[] buffer = new byte[4096];
            int bytesLidos;

            while ((bytesLidos = entrada.read(buffer)) != -1) {
                saida.write(buffer, 0, bytesLidos);
            }
            saida.close();
            entrada.close();
            System.out.println("✅ Download concluído: " + nomeDestino);
        }catch (Exception e){
                System.out.println("⚠️ Falha ao descarregar o ficheiro: " + nomeDestino);

        }
    }
        
    public static void fazerComparação(String caminho1, String caminho2, List<Integer> ignorarColunas, boolean ignorarMaiusculas){
        try(FileInputStream fis1 = new FileInputStream(new File(caminho1));
            FileInputStream fis2 = new FileInputStream(new File(caminho2));
            Workbook wb1 = WorkbookFactory.create(fis1);
            Workbook wb2 = WorkbookFactory.create(fis2);
            FileWriter fw = new FileWriter("C:/Users/AEB_Aluno/Downloads/relatorio.txt");
            PrintWriter escritor = new PrintWriter(fw)) {

            int TotalDeLinhasProcessadas = 0;
            int TotalDeDiferençasEncontradas = 0;  

            Sheet folha1 = wb1.getSheetAt(0);
            Sheet folha2 = wb2.getSheetAt(0);

            Iterator<Row> iteratorFolha1 = folha1.rowIterator();
            Iterator<Row> iteratorFolha2 = folha2.rowIterator();

            int ContadorLinhas = 1;
            boolean EncontrouDiferençaNoTotal = false;

            while(iteratorFolha1.hasNext() || iteratorFolha2.hasNext()){
                System.out.print("\r" + AMARELO + "A processar linhas" + ContadorLinhas + RESET);

                Row linha1 = iteratorFolha1.hasNext() ? iteratorFolha1.next(): null;
                Row linha2 = iteratorFolha2.hasNext() ? iteratorFolha2.next(): null;

                ContadorLinhas++;
                
                int maxColunas = 0;
                if (linha1 != null) maxColunas = Math.max(maxColunas, linha1.getLastCellNum());
                if (linha2 != null) maxColunas = Math.max(maxColunas, linha2.getLastCellNum());

                for(int c =0; c < maxColunas; c++){
                    if(ignorarColunas.contains(c)){
                        continue;
                    }

                    String valor1 = obtertextoCelula(linha1, c);
                    String valor2 = obtertextoCelula(linha2, c);

                    boolean saoIguais;
                    if(ignorarMaiusculas){
                        saoIguais = valor1.equalsIgnoreCase(valor2);
                    }else{
                        saoIguais = valor1.equals(valor2);
                    }

                    if(!saoIguais){
                        EncontrouDiferençaNoTotal = true;
                        String LetraColuna = CellReference.convertNumToColString(c);

                        System.out.printf("DETETADA DIFERENÇA (coluna %s): %n", LetraColuna);
                        escritor.printf("DEFERENÇA DETETADA na linha %d... %n", (ContadorLinhas));
                        System.out.printf("   BOM_FO - Linha %d - Valor: [%s]%n", (ContadorLinhas), valor1);
                        System.out.printf("   BOM_MES - Linha %d - Valor: [%s]%n", (ContadorLinhas), valor2);
                        System.out.println("   ------------------------------------------------");
                        TotalDeDiferençasEncontradas++;
                    }
                }   
                TotalDeLinhasProcessadas++;
            }

            if(EncontrouDiferençaNoTotal){
                System.out.println("\r");
                System.out.println(CIANO + "=================================================" + RESET);
                System.out.println(CIANO + "       RESUNO AUDITORIAL DE PROCESSAMENTO        " + RESET);
                System.out.println(CIANO + "=================================================" + RESET);
                System.out.printf("  -> Total de Linhas Analisadas: %d%n", TotalDeLinhasProcessadas);

                if(TotalDeDiferençasEncontradas > 0){
                    System.out.println(VERMELHO + "STATUS : DEPENDENCIAS DETETADAS" + RESET);
                    System.out.printf(VERMELHO + "Total de conflitos Encontrados %d%n" + RESET, TotalDeDiferençasEncontradas);
                    System.out.printf("       FIM DO PROCESSAMENTO - TOTAL DE CONFLITOS: %d%n       ", TotalDeDiferençasEncontradas);
                }else{
                    System.out.println(VERDE + "FICHEIRO EM COMFORMIDADE (100% IGUAIS)" + RESET);
                    escritor.println("Auditoria Concluída: Nenhuma Descrepância encontrada.");
                }
                System.out.println(CIANO + "=========================================================" + RESET);

                if(TotalDeDiferençasEncontradas > 0){
                    System.out.println(VERDE + "O relatório detalhado foi exportado com sucesso para a pasta Downloads."+ RESET);
                }
            }
        }catch(Exception e){
            System.out.println("Erro ao ler os ficheiros Excel no motor de comparação:");
            e.printStackTrace();
        }
    }

    public static void fazerComparaçãoCSV(String Caminho1, String Caminho2, List<Integer> ignorarColunas, boolean ignorarMaiusculas){
        String separadores = DescobrirSeparador(Caminho1);
        System.out.println(AMARELO + "Separador detectado automaticamente: [" + separadores + "]" + RESET);

        try(BufferedReader br1 = new BufferedReader(new FileReader(Caminho1));
            BufferedReader br2 = new BufferedReader(new FileReader(Caminho2));
            FileWriter fw = new FileWriter("C:/Users/AEB_Aluno/Downloads/relatorio_csv.txt");
            PrintWriter escritor = new PrintWriter(fw)){

            int TotalDeLinhasProcessadas = 0;
            int TotalDeDiferençasEncontradas = 0;  
            
            int ContadorLinhas = 1;
            boolean EncontrouDiferençaNoTotal = false;

            escritor.println("=====================================================");
            escritor.println("RELATÓRIOS DE DIFERENÇAS DETETADAS (FORMATO CSV)");
            escritor.println("=====================================================");

            String linha1, linha2;

            while((linha1 = br1.readLine()) != null | (linha2 = br2.readLine()) != null){
                System.out.print("\r" + AMARELO + "A processar linha CSV: " + ContadorLinhas + RESET);

                String texto1 = (linha1 != null ) ? linha1: "";
                String texto2 = (linha2 != null ) ? linha2: "";

                String Regex ="(?=(?:[^\"]*\"[^\"]*\")*[^\"]*$)" + separadores;

                String[] colunas1 = texto1.split(Regex, -1);
                String[] colunas2 = texto2.split(Regex, -1);

                int maxColunas = Math.max(colunas1.length, colunas2.length);

                for (int c  = 0; c < maxColunas; c++){
                    if(ignorarColunas.contains(c)) continue;
                    String valor1 = (c < colunas1.length) ? colunas1[c].replace("\"", "").trim() : "VAZIO";
                    String valor2 = (c < colunas2.length) ? colunas2[c].replace("\"", "").trim() : "VAZIO";

                    boolean saoIguais = ignorarMaiusculas ? valor1.equalsIgnoreCase(valor2) : valor1.equals(valor2);

                    if(!saoIguais){
                        EncontrouDiferençaNoTotal = true;
                        System.out.printf("%nDIFERENÇA DETETADA EM CSV (coliuna índice %d) : %n", c);
                        System.out.printf("BOM_FO - Linha %d - Valor: [%s]%n", ContadorLinhas, valor1);
                        System.out.printf("BOM_MES - Linha %d - Valor: [%s]%n", ContadorLinhas, valor2);

                        escritor.printf("DIFERENÇA DETECTADA (coluna indice %d) na linha %d... %n", c, ContadorLinhas);
                        TotalDeDiferençasEncontradas++;
                    }
                }
                ContadorLinhas++;
                TotalDeLinhasProcessadas++;
            }

            System.out.print("\r");
            System.out.println(CIANO + "=================================================" + RESET);
            System.out.println(CIANO + "       RESUMO AUDITORIAL DE PROCESSAMENTO (CSV)  " + RESET);
            System.out.println(CIANO + "=================================================" + RESET);
            System.out.printf("  -> Total de Linhas Analisadas: %d%n", TotalDeLinhasProcessadas);
            
            if (EncontrouDiferençaNoTotal) {
                System.out.println(VERMELHO + "  -> STATUS: CONFLITOS DETETADOS" + RESET);
                System.out.printf(VERMELHO + "  -> Total de Discrepâncias: %d%n" + RESET, TotalDeDiferençasEncontradas);
                escritor.printf("%nFIM DO PROCESSAMENTO - TOTAL DE CONFLITOS: %d%n", TotalDeDiferençasEncontradas);
            } else {
                System.out.println(VERDE + "  -> STATUS: FICHEIROS EM CONFORIDADE (100% IGUAIS)" + RESET);
                escritor.println("Auditoria Concluída: Nenhuma discrepância encontrada.");
            }
            System.out.println(CIANO + "=================================================" + RESET);

        }catch(Exception e){
            System.out.println(VERMELHO + "Erro ao ler os ficheiros CSV no motor de comparação." + RESET);
            e.printStackTrace();
        }
    }

    private static String obtertextoCelula(Row linha, int indiceColuna){
        if(linha == null) return "LINHA VAZIA" ;
        Cell celula = linha.getCell(indiceColuna);
        if(celula == null || celula.getCellType() == CellType.BLANK) return "VAZIO";
        
        switch (celula.getCellType()){
            case STRING:
                return celula.getStringCellValue().trim();
            case NUMERIC:
                if (DateUtil.isCellDateFormatted(celula)){
                    return celula.getDateCellValue().toString();
                }
                double num = celula.getNumericCellValue();
                if(num == (long) num){
                    return String.format("%d", (long) num);
                }else{
                    return String.valueOf(num);
                }
            case BOOLEAN:
                return String.valueOf(celula.getBooleanCellValue());
            default:
                return "CONTÉUDO NÃO ENCONTRADO";       
        }  
    }

    private static void mostrarMenuColunas(String caminhoFicheiro) {
        try(FileInputStream fis1 = new FileInputStream(new File(caminhoFicheiro))){
            Workbook wb = WorkbookFactory.create(fis1);
            Sheet folha = wb.getSheetAt(0);
            Row PrimeiraLinha = folha.getRow(0);

            System.out.println(AMARELO + "\n ----- COLUNAS DETETADAS NO DOCUMENTO -----" + RESET);
            if(PrimeiraLinha != null){
                for (int c =0; c < PrimeiraLinha.getLastCellNum(); c++){
                    Cell celula = PrimeiraLinha.getCell(c);
                    String nomeColuna = (celula != null) ? celula.getStringCellValue() : "";
                    System.out.printf("[%d] -> %s%n", c, nomeColuna);
                }
            }
        }catch(Exception e){
            System.out.println("Erro na leitura dos cabeçalhos");
        } 
    }
        
    public static String DescobrirSeparador(String caminho){
        try (BufferedReader br = new BufferedReader(new FileReader(caminho))){
            String PrimeiraLinha = br.readLine();
            if(PrimeiraLinha != null){
                long virgula = PrimeiraLinha.chars().filter(ch -> ch == ',').count();
                long PontosVirgula = PrimeiraLinha.chars().filter(ch -> ch == ';').count();

                return (PontosVirgula >= virgula) ? ";":",";
            }
        }catch(Exception e){
            // Erro silencioso tratado
        }
        return ";"; 
    } 
     
    private static int ObterTotalColunasExcel(String caminho){
        try(FileInputStream fis = new FileInputStream(new File(caminho));
            Workbook wb = WorkbookFactory.create(fis)){
            Sheet folha = wb.getSheetAt(0);
            Row PrimeiraLinha = folha.getRow(0);
            return (PrimeiraLinha != null) ? PrimeiraLinha.getLastCellNum() : 0;
        }catch(Exception e){
            return 0;
        }
    }

    private static int ObterTotalColunasCSV(String caminho, String separador){
        try(BufferedReader br = new BufferedReader(new FileReader(caminho))){
            String PrimeiraLinha = br.readLine();
            if(PrimeiraLinha != null){
                String regex = "(?=(?:[^\"]*\"[^\"]*\")*[^\"]*$)" + separador;
                String[] colunas = PrimeiraLinha.split(regex, -1);
                return colunas.length;
            }
        }catch(Exception e){
            // Erro tratado
        }
        return 0;
    }
}