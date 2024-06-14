package Del;

import javax.swing.*;
import javax.swing.event.MenuListener;

import java.awt.*;
import java.awt.event.ActionEvent;
import java.awt.event.ActionListener;
import java.awt.image.BufferedImage;

public class ExcelHome extends JFrame {

    private JTextArea textArea;
    private ImagePanel imagePanel;
    
    private JButton item1;
    private JButton item2;
    private JButton item3;
    private JButton item4;
    private JButton item5;

    public ExcelHome() {
        setTitle("배드민턴 통합 정보");
        setSize(800, 600);
        setDefaultCloseOperation(JFrame.EXIT_ON_CLOSE);
        setLayout(null); // 레이아웃을 null로 설정하여 수동 배치

        MyPanel temper = new MyPanel();
        temper.setBounds(400, 0, 300, 100);
        add(temper);
        temper.setVisible(true);
        createMenu();
        
        
        // 버튼 생성 및 필드로 초기화
        item1 = new JButton("Item1");
        item2 = new JButton("Item2");
        item3 = new JButton("Item3");
        item4 = new JButton("Item4");
        item5 = new JButton("Item5");
        item1.setBounds(100, 100, 75, 25);
        item2.setBounds(175, 100, 75, 25);
        item3.setBounds(250, 100, 75, 25);
        item4.setBounds(325, 100, 75, 25);
        item5.setBounds(400, 100, 75, 25);

        // 프레임에 버튼 추가
        add(item1);
        add(item2);
        add(item3);
        add(item4);
        add(item5);
        item1.setVisible(false);
        item2.setVisible(false);
        item3.setVisible(false);
        item4.setVisible(false);
        item5.setVisible(false);
        
        // 텍스트 영역 생성
        textArea = new JTextArea();
        textArea.setEditable(false);
        textArea.setFont(new Font("Monospaced", Font.PLAIN, 10));
        JScrollPane scrollPane = new JScrollPane(textArea);
        scrollPane.setBounds(400, 150, 200, 200); // 위치와 크기 설정

        // 이미지를 그릴 패널 생성
        imagePanel = new ImagePanel();
        imagePanel.setBounds(100, 150, 200, 200); // 위치와 크기 설정

        // 컴포넌트를 프레임에 추가
        add(scrollPane);
        add(imagePanel);

        setVisible(true);
    }

    private void createMenu() {//MenuBar에 위치한 옵션들을 생성 및 초기화
        JMenuBar mb = new JMenuBar();
        JMenu screenMenu = new JMenu("Browsing");
        screenMenu.addSeparator(); // 분리선 삽입

        JMenuItem racketMenuItem = new JMenuItem("Racket");
        racketMenuItem.addActionListener(new MenuActionListener());

        JMenuItem shoesMenuItem = new JMenuItem("Shoes");
        shoesMenuItem.addActionListener(new MenuActionListener());

        JMenuItem bagMenuItem = new JMenuItem("Bag");
        bagMenuItem.addActionListener(new MenuActionListener());

        JMenuItem shuttleCockMenuItem = new JMenuItem("Shuttle Cock");
        shuttleCockMenuItem.addActionListener(new MenuActionListener());
        
        JMenuItem rulesMenuItem = new JMenuItem("Rules");
        rulesMenuItem.addActionListener(new ActionListener() {
            @Override
            public void actionPerformed(ActionEvent e) {
                RulesloadSheetData(0);
            }
        });
        
        screenMenu.add(racketMenuItem);
        screenMenu.add(shoesMenuItem);
        screenMenu.add(bagMenuItem);
        screenMenu.add(shuttleCockMenuItem);
        screenMenu.add(rulesMenuItem);
        
        mb.add(screenMenu);
        setJMenuBar(mb);
    }

    public void RacketloadSheetData(int sheetIndex) {
        RacketDataLoader loader = new RacketDataLoader(textArea, imagePanel);
        loader.loadExcelData(sheetIndex);
    }

    public void ShoesloadSheetData(int sheetIndex) {
        ShoesDataLoader loader = new ShoesDataLoader(textArea, imagePanel);
        loader.loadExcelData(sheetIndex);
    }

    public void BagloadSheetData(int sheetIndex) {
        BagDataLoader loader = new BagDataLoader(textArea, imagePanel);
        loader.loadExcelData(sheetIndex);
    }

    public void ShuttleCockloadSheetData(int sheetIndex) {
        ShuttleCockDataLoader loader = new ShuttleCockDataLoader(textArea, imagePanel);
        loader.loadExcelData(sheetIndex);
    }
    
    public void RulesloadSheetData(int sheetIndex) {
        RulesDataLoader loader = new RulesDataLoader(textArea);
        loader.loadExcelData(sheetIndex);
    }

    // 메뉴 아이템 선택에 따른 결과
    class MenuActionListener implements ActionListener {
        public void actionPerformed(ActionEvent e) {
            String cmd = e.getActionCommand();
            Component[] components = getContentPane().getComponents();

            if ("Racket".equals(cmd)) {
                // Racket 메뉴 선택 시 버튼들을 보이게 설정
                for (Component component : components) {
                    if (component instanceof JButton) {
                        component.setVisible(true);
                    }
                }
                RacketloadSheetData(0); // 기본적으로 첫 번째 시트를 로드

                // 기존 리스너 제거
                for (ActionListener al : item1.getActionListeners()) item1.removeActionListener(al);
                for (ActionListener al : item2.getActionListeners()) item2.removeActionListener(al);
                for (ActionListener al : item3.getActionListeners()) item3.removeActionListener(al);
                for (ActionListener al : item4.getActionListeners()) item4.removeActionListener(al);
                for (ActionListener al : item5.getActionListeners()) item5.removeActionListener(al);

                // 새로운 리스너 설정
                item1.addActionListener(e1 -> RacketloadSheetData(0));
                item2.addActionListener(e1 -> RacketloadSheetData(1));
                item3.addActionListener(e1 -> RacketloadSheetData(2));
                item4.addActionListener(e1 -> RacketloadSheetData(3));
                item5.addActionListener(e1 -> RacketloadSheetData(4));
            } else if ("Shoes".equals(cmd)) {
                // Shoes 메뉴 선택 시 버튼들을 보이게 설정
                for (Component component : components) {
                    if (component instanceof JButton) {
                        component.setVisible(true);
                    }
                }
                ShoesloadSheetData(0); // 기본적으로 첫 번째 시트를 로드

                // 기존 리스너 제거
                for (ActionListener al : item1.getActionListeners()) item1.removeActionListener(al);
                for (ActionListener al : item2.getActionListeners()) item2.removeActionListener(al);
                for (ActionListener al : item3.getActionListeners()) item3.removeActionListener(al);
                for (ActionListener al : item4.getActionListeners()) item4.removeActionListener(al);
                for (ActionListener al : item5.getActionListeners()) item5.removeActionListener(al);

                // 새로운 리스너 설정
                item1.addActionListener(e1 -> ShoesloadSheetData(0));
                item2.addActionListener(e1 -> ShoesloadSheetData(1));
                item3.addActionListener(e1 -> ShoesloadSheetData(2));
                item4.addActionListener(e1 -> ShoesloadSheetData(3));
                item5.addActionListener(e1 -> ShoesloadSheetData(4));
            } else if ("Bag".equals(cmd)) {
                // Bag 메뉴 선택 시 버튼들을 보이게 설정
                for (Component component : components) {
                    if (component instanceof JButton) {
                        component.setVisible(true);
                    }
                }
                BagloadSheetData(0); // 기본적으로 첫 번째 시트를 로드

                // 기존 리스너 제거
                for (ActionListener al : item1.getActionListeners()) item1.removeActionListener(al);
                for (ActionListener al : item2.getActionListeners()) item2.removeActionListener(al);
                for (ActionListener al : item3.getActionListeners()) item3.removeActionListener(al);
                for (ActionListener al : item4.getActionListeners()) item4.removeActionListener(al);
                for (ActionListener al : item5.getActionListeners()) item5.removeActionListener(al);

                // 새로운 리스너 설정
                item1.addActionListener(e1 -> BagloadSheetData(0));
                item2.addActionListener(e1 -> BagloadSheetData(1));
                item3.addActionListener(e1 -> BagloadSheetData(2));
                item4.addActionListener(e1 -> BagloadSheetData(3));
                item5.addActionListener(e1 -> BagloadSheetData(4));
            } else if ("Shuttle Cock".equals(cmd)) {
                // Shuttle Cock 메뉴 선택 시 버튼들을 보이게 설정
                for (Component component : components) {
                    if (component instanceof JButton) {
                        component.setVisible(true);
                    }
                }
                ShuttleCockloadSheetData(0); // 기본적으로 첫 번째 시트를 로드

                // 기존 리스너 제거
                for (ActionListener al : item1.getActionListeners()) item1.removeActionListener(al);
                for (ActionListener al : item2.getActionListeners()) item2.removeActionListener(al);
                for (ActionListener al : item3.getActionListeners()) item3.removeActionListener(al);
                for (ActionListener al : item4.getActionListeners()) item4.removeActionListener(al);
                for (ActionListener al : item5.getActionListeners()) item5.removeActionListener(al);

                // 새로운 리스너 설정
                item1.addActionListener(e1 -> ShuttleCockloadSheetData(0));
                item2.addActionListener(e1 -> ShuttleCockloadSheetData(1));
                item3.addActionListener(e1 -> ShuttleCockloadSheetData(2));
                item4.addActionListener(e1 -> ShuttleCockloadSheetData(3));
                item5.addActionListener(e1 -> ShuttleCockloadSheetData(4));
            }
        }
    }

    class MyPanel extends JPanel {
        public void paintComponent(Graphics g) {
            super.paintComponent(g);

            g.setColor(new Color(0, 0, 0)); // 검은색 지정
            g.setFont(new Font("Arial", Font.ITALIC, 20));
            g.drawString("Badminton Integrated System", 30, 70);

        }
    }

    public static void main(String[] args) {
        // UI를 이벤트 디스패치 스레드에서 실행
        SwingUtilities.invokeLater(ExcelHome::new);
    }
}

// 액셀에서 이미지를 읽은 것을 스윙 프레임에 그리기 위해서 무조건 필요.
class ImagePanel extends JPanel { 
    private BufferedImage img;

    // 이미지를 설정하는 메서드
    public void setImage(BufferedImage img) {
        this.img = img;
        repaint(); // 패널을 다시 그리도록 요청
    }

    @Override
    protected void paintComponent(Graphics g) {
        super.paintComponent(g);
        if (img != null) {
            // 이미지가 패널 크기에 맞게 조정되도록 그림.
            g.drawImage(img, 0, 0, getWidth(), getHeight(), this);
        }
    }
}
