import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.DateUtil;
import org.apache.poi.ss.util.CellReference;
import java.util.List;
import java.util.ArrayList;
import java.util.Iterator;
import javax.net.ssl.HttpsURLConnection;
import java.net.URL;
import java.net.HttpURLConnection;
import java.io.*;
import javax.swing.*;
import javax.swing.border.LineBorder;
import java.awt.*;

public class Main {

    public static void main(String[] args) {
        // 🔥 Ativar o Nimbus Modo Dark (Nativo do Swing, sem bloqueios da escola)
        try {
            for (UIManager.LookAndFeelInfo info : UIManager.getInstalledLookAndFeels()) {
                if ("Nimbus".equals(info.getName())) {
                    UIManager.setLookAndFeel(info.getClassName());
                    break;
                }
            }
            // Customizar as cores do Nimbus para ficar escuro futurista
            UIManager.put("nimbusBase", new Color(18, 18, 24));      
            UIManager.put("nimbusBlueGrey", new Color(30, 30, 40));  
            UIManager.put("control", new Color(18, 18, 24));         
            UIManager.put("text", new Color(240, 243, 246));         
            UIManager.put("nimbusLightBackground", new Color(30, 30, 40));
            UIManager.put("info", new Color(30, 30, 40));
        } catch (Exception e) {
            System.out.println("Nao foi possivel aplicar o tema.");
        }

        String versaoLocal = "1.1";
        System.out.println("A verificar existencia de atualizacoes na nuvem...");
        String versaoNuvem = ObterVersaoNuvem();
        
        if(versaoNuvem != null){
            if(!versaoLocal.equals(versaoNuvem)){
                System.out.println("Nova versao existente: (Nuvem: " + versaoNuvem + " | Local: " + versaoLocal + ")");
                System.out.println("A iniciar download da nova versao...");

                String urlBase = "https://raw.githubusercontent.com/anonimoani31-rgb/Comparador-Updates/refs/heads/main/";
                descarregarAtualizacao(urlBase + "Main.class", "Main.class");
                descarregarAtualizacao(urlBase + "Main$1.class", "Main$1.class");
                System.out.println("Programa updated com exito! Por favor reinicie o programa.");
            }else{
                System.out.println("O programa esta atualizado para a versao: " + versaoLocal);
            }
        }
        
        SwingUtilities.invokeLater(() -> criarInterfaceGrafica());
    }

   public static void criarInterfaceGrafica() {
        Color azulNeon = new Color(0, 180, 216);         
        Color verdeNeon = new Color(0, 245, 212);         
        Color textoBranco = new Color(240, 243, 246);
        
        JFrame janela = new JFrame("Comparador Avancado de Ficheiros");
        janela.setSize(600, 420); // Aumentámos um pouco a altura para caber o novo campo
        janela.setDefaultCloseOperation(JFrame.EXIT_ON_CLOSE);
        janela.setLocationRelativeTo(null);
        janela.setLayout(null);

        // Ficheiro 1
        JLabel lblFicheiro1 = new JLabel("Ficheiro 1:");
        lblFicheiro1.setBounds(30, 40, 80, 25);
        lblFicheiro1.setFont(new Font("Segoe UI", Font.BOLD, 12));
        lblFicheiro1.setForeground(azulNeon); 
        janela.add(lblFicheiro1);

        JTextField txtFicheiro1 = new JTextField();
        txtFicheiro1.setBounds(120, 40, 300, 26);
        txtFicheiro1.setForeground(textoBranco);
        txtFicheiro1.setCaretColor(azulNeon); 
        txtFicheiro1.setFont(new Font("Segoe UI", Font.PLAIN, 12));
        txtFicheiro1.setBorder(new LineBorder(azulNeon, 1));
        janela.add(txtFicheiro1);

        JButton btnProcurar1 = new JButton("Procurar...");
        btnProcurar1.setBounds(440, 40, 110, 26);
        btnProcurar1.setForeground(azulNeon);
        btnProcurar1.setFont(new Font("Segoe UI", Font.BOLD, 12));
        btnProcurar1.setBorderPainted(false);
        btnProcurar1.setContentAreaFilled(false);
        btnProcurar1.setFocusPainted(false);
        janela.add(btnProcurar1);

        // Ficheiro 2
        JLabel lblFicheiro2 = new JLabel("Ficheiro 2:");
        lblFicheiro2.setBounds(30, 100, 80, 25);
        lblFicheiro2.setFont(new Font("Segoe UI", Font.BOLD, 12));
        lblFicheiro2.setForeground(azulNeon);
        janela.add(lblFicheiro2);

        JTextField txtFicheiro2 = new JTextField();
        txtFicheiro2.setBounds(120, 100, 300, 26);
        txtFicheiro2.setForeground(textoBranco);
        txtFicheiro2.setCaretColor(azulNeon);
        txtFicheiro2.setFont(new Font("Segoe UI", Font.PLAIN, 12));
        txtFicheiro2.setBorder(new LineBorder(azulNeon, 1));
        janela.add(txtFicheiro2);

        JButton btnProcurar2 = new JButton("Procurar...");
        btnProcurar2.setBounds(440, 100, 110, 26);
        btnProcurar2.setForeground(azulNeon);
        btnProcurar2.setFont(new Font("Segoe UI", Font.BOLD, 12));
        btnProcurar2.setBorderPainted(false);
        btnProcurar2.setContentAreaFilled(false);
        btnProcurar2.setFocusPainted(false);
        janela.add(btnProcurar2);

        // 🔥 NOVO: Campo para Ignorar Colunas
        JLabel lblIgnorarColunas = new JLabel("Ignorar Cols (Ex: 0,2):");
        lblIgnorarColunas.setBounds(30, 160, 140, 25);
        lblIgnorarColunas.setFont(new Font("Segoe UI", Font.BOLD, 12));
        lblIgnorarColunas.setForeground(azulNeon);
        janela.add(lblIgnorarColunas);

        JTextField txtIgnorarColunas = new JTextField();
        txtIgnorarColunas.setBounds(180, 160, 240, 26);
        txtIgnorarColunas.setForeground(textoBranco);
        txtIgnorarColunas.setCaretColor(azulNeon);
        txtIgnorarColunas.setFont(new Font("Segoe UI", Font.PLAIN, 12));
        txtIgnorarColunas.setBorder(new LineBorder(azulNeon, 1));
        janela.add(txtIgnorarColunas);

        // Checkbox Maiúsculas/Minúsculas (Empurrada um pouco para baixo)
        JCheckBox chkIgnorar = new JCheckBox("Ignorar Maiusculas/Minusculas");
        chkIgnorar.setBounds(180, 210, 250, 25);
        chkIgnorar.setFont(new Font("Segoe UI", Font.PLAIN, 12));
        janela.add(chkIgnorar);

        // Botão Comparar
        JButton btnComparar = new JButton("Iniciar Comparacao");
        btnComparar.setBounds(200, 280, 200, 45);
        btnComparar.setForeground(verdeNeon);
        btnComparar.setFont(new Font("Segoe UI", Font.BOLD, 14)); 
        btnComparar.setBorderPainted(false);
        btnComparar.setContentAreaFilled(false);
        btnComparar.setFocusPainted(false);
        janela.add(btnComparar);

        // Ações dos Botões Procurar
        btnProcurar1.addActionListener(e -> {
            JFileChooser seletor = new JFileChooser();
            if (seletor.showOpenDialog(janela) == JFileChooser.APPROVE_OPTION) {
                txtFicheiro1.setText(seletor.getSelectedFile().getAbsolutePath());
            }
        });

        btnProcurar2.addActionListener(e -> {
            JFileChooser seletor = new JFileChooser();
            if (seletor.showOpenDialog(janela) == JFileChooser.APPROVE_OPTION) {
                txtFicheiro2.setText(seletor.getSelectedFile().getAbsolutePath());
            }
        });

        // Ação do Botão Comparar
        btnComparar.addActionListener(e -> {
            String caminho1 = txtFicheiro1.getText().trim();
            String caminho2 = txtFicheiro2.getText().trim();
            boolean ignorar = chkIgnorar.isSelected();

            if (caminho1.isEmpty() || caminho2.isEmpty()) {
                JOptionPane.showMessageDialog(janela, "Por favor, selecione ambos os ficheiros!", "Erro", JOptionPane.ERROR_MESSAGE);
                return;
            }

            File f1 = new File(caminho1);
            File f2 = new File(caminho2);

            if (!f1.exists() || !f2.exists()) {
                JOptionPane.showMessageDialog(janela, "Ficheiros nao encontrados!", "Erro Critico", JOptionPane.ERROR_MESSAGE);
                return;
            }

            String extF1 = caminho1.substring(caminho1.lastIndexOf(".")).toLowerCase();
            String extF2 = caminho2.substring(caminho2.lastIndexOf(".")).toLowerCase();

            if(!extF1.equals(extF2)){
                JOptionPane.showMessageDialog(janela, "Formatos diferentes!", "Erro de Formato", JOptionPane.ERROR_MESSAGE);
                return;
            }

            // 🔥 PROCESSAR AS COLUNAS A IGNORAR
            List<Integer> colunasParaIgnorar = new ArrayList<>();
            String textoColunas = txtIgnorarColunas.getText().trim();
            if (!textoColunas.isEmpty()) {
                String[] partes = textoColunas.split(",");
                for (String parte : partes) {
                    try {
                        colunasParaIgnorar.add(Integer.parseInt(parte.trim()));
                    } catch (NumberFormatException nfe) {
                        // Se o utilizador meter uma letra sem querer, o programa ignora e não rebenta
                    }
                }
            }

            int numColunasF1 = 0;
            int numColunasF2 = 0;

            if(extF1.equals(".xlsx") || extF1.equals(".xls")){
                numColunasF1 = ObterTotalColunasExcel(caminho1);
                numColunasF2 = ObterTotalColunasExcel(caminho2);
            } else if(extF1.equals(".csv")){
                String sep = DescobrirSeparador(caminho1);
                numColunasF1 = ObterTotalColunasCSV(caminho1, sep);
                numColunasF2 = ObterTotalColunasCSV(caminho2, sep);
            }

            if (numColunasF1 != numColunasF2) {
                int resposta = JOptionPane.showConfirmDialog(janela, 
                    "Estruturas de colunas diferentes!\nFicheiro 1: " + numColunasF1 + " | Ficheiro 2: " + numColunasF2 + "\nForcar comparacao?", 
                    "Aviso", JOptionPane.YES_NO_OPTION);
                if (resposta != JOptionPane.YES_OPTION) return;
            }

            janela.setVisible(false);
            
            // 🔥 Passar a lista real de colunas ignoradas para os métodos
            if(extF1.equals(".csv")) {
                fazerComparacaoCSV(caminho1, caminho2, colunasParaIgnorar, ignorar);
            } else {
                fazerComparacao(caminho1, caminho2, colunasParaIgnorar, ignorar);
            }
            
            JOptionPane.showMessageDialog(null, "Comparacao concluida! Verifica os logs.", "Sucesso", JOptionPane.INFORMATION_MESSAGE);
            System.exit(0);
        });

        janela.setVisible(true);
    }

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
        }catch (Exception e){}
    }

    public static void fazerComparacao(String caminho1, String caminho2, List<Integer> ignorarColunas, boolean ignorarMaiusculas){
        Workbook wb1 = null;  Workbook wb2 = null;
        try(FileInputStream fis1 = new FileInputStream(new File(caminho1));
            FileInputStream fis2 = new FileInputStream(new File(caminho2));
            FileWriter fw = new FileWriter("C:/Users/AEB_Aluno/Downloads/relatorio.txt");
            PrintWriter escritor = new PrintWriter(fw)) {
            
            wb1 = WorkbookFactory.create(fis1); 
            wb2 = WorkbookFactory.create(fis2); 
            
            Sheet folha1 = wb1.getSheetAt(0); 
            Sheet folha2 = wb2.getSheetAt(0);
            
            // 🔥 Encontrar o número máximo real de linhas entre os dois ficheiros
            int maxLinhas = Math.max(folha1.getLastRowNum(), folha2.getLastRowNum());
            
            // 🔥 Percorrer linha a linha por índice fixo (Garante que a linha bate sempre certo!)
            for (int i = 0; i <= maxLinhas; i++) {
                Row linha1 = folha1.getRow(i);
                Row linha2 = folha2.getRow(i);
                
                // O número da linha para o Excel é o índice do Java + 1
                int numeroLinhaExcel = i + 1; 

                // Se ambas as linhas não existem no ficheiro, saltamos
                if (linha1 == null && linha2 == null) continue;

                int maxColunas = 0;
                if (linha1 != null) maxColunas = Math.max(maxColunas, linha1.getLastCellNum());
                if (linha2 != null) maxColunas = Math.max(maxColunas, linha2.getLastCellNum());
                
                for(int c = 0; c < maxColunas; c++){
                    if(ignorarColunas.contains(c)) continue;
                    
                    String valor1 = obtertextoCelula(linha1, c); 
                    String valor2 = obtertextoCelula(linha2, c);
                    
                    boolean saoIguais = ignorarMaiusculas ? valor1.equalsIgnoreCase(valor2) : valor1.equals(valor2);
                    
                    if(!saoIguais){
                        String LetraColuna = CellReference.convertNumToColString(c);
                        escritor.printf("DIFERENCA DETETADA na linha %d, coluna %s... (F1: [%s] vs F2: [%s])%n", 
                                        numeroLinhaExcel, LetraColuna, valor1, valor2);
                    }
                }   
            }
        } catch(Exception e) {
            e.printStackTrace();
        } finally {
            try { if (wb1 != null) wb1.close(); if (wb2 != null) wb2.close(); } catch (Exception ex) {}
        }
    }

    public static void fazerComparacaoCSV(String Caminho1, String Caminho2, List<Integer> ignorarColunas, boolean ignorarMaiusculas){
        String separadores = DescobrirSeparador(Caminho1);
        try(BufferedReader br1 = new BufferedReader(new FileReader(Caminho1));
            BufferedReader br2 = new BufferedReader(new FileReader(Caminho2));
            FileWriter fw = new FileWriter("C:/Users/AEB_Aluno/Downloads/relatorio_csv.txt");
            PrintWriter escritor = new PrintWriter(fw)){
            int ContadorLinhas = 1;
            while(true) {
                String linha1 = br1.readLine(); String linha2 = br2.readLine();
                if (linha1 == null && linha2 == null) break;
                String texto1 = (linha1 != null) ? linha1 : ""; String texto2 = (linha2 != null) ? linha2 : "";
                String Regex = "(?=(?:[^\"]*\"[^\"]*\")*[^\"]*$)" + separadores;
                String[] colunas1 = texto1.split(Regex, -1); String[] colunas2 = texto2.split(Regex, -1);
                int maxColunas = Math.max(colunas1.length, colunas2.length);
                for (int c = 0; c < maxColunas; c++){
                    if(ignorarColunas.contains(c)) continue;
                    String valor1 = (c < colunas1.length) ? colunas1[c].replace("\"", "").trim() : "VAZIO";
                    String valor2 = (c < colunas2.length) ? colunas2[c].replace("\"", "").trim() : "VAZIO";
                    boolean saoIguais = ignorarMaiusculas ? valor1.equalsIgnoreCase(valor2) : valor1.equals(valor2);
                    if(!saoIguais){
                        escritor.printf("DIFERENCA CSV: Linha %d, Col %d [%s] vs [%s]%n", ContadorLinhas, c, valor1, valor2);
                    }
                }
                ContadorLinhas++;
            }
        }catch(Exception e){}
    }

    private static String obtertextoCelula(Row linha, int indiceColuna){
        if(linha == null) return "LINHA VAZIA" ;
        Cell celula = linha.getCell(indiceColuna);
        if(celula == null || celula.getCellType() == CellType.BLANK) return "VAZIO";
        switch (celula.getCellType()){
            case STRING: return celula.getStringCellValue().trim();
            case NUMERIC:
                if (DateUtil.isCellDateFormatted(celula)) return celula.getDateCellValue().toString();
                double num = celula.getNumericCellValue();
                return (num == (long) num) ? String.format("%d", (long) num) : String.valueOf(num);
            case BOOLEAN: return String.valueOf(celula.getBooleanCellValue());
            default: return "CONTEUDO NAO ENCONTRADO";       
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
        }catch(Exception e){}
        return ";"; 
    }

    private static int ObterTotalColunasExcel(String caminho){
        try(FileInputStream fis = new FileInputStream(new File(caminho));
            Workbook wb = WorkbookFactory.create(fis)){
            Sheet folha = wb.getSheetAt(0); Row PrimeiraLinha = folha.getRow(0);
            return (PrimeiraLinha != null) ? PrimeiraLinha.getLastCellNum() : 0;
        }catch(Exception e){ return 0; }
    }

    private static int ObterTotalColunasCSV(String caminho, String separador){
        try(BufferedReader br = new BufferedReader(new FileReader(caminho))){
            String PrimeiraLinha = br.readLine();
            if(PrimeiraLinha != null){
                String regex = "(?=(?:[^\"]*\"[^\"]*\")*[^\"]*$)" + separador;
                return PrimeiraLinha.split(regex, -1).length;
            }
        }catch(Exception e){}
        return 0;
    }
}