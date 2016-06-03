package lms.swing;

import java.awt.Cursor;
import java.awt.EventQueue;
import java.awt.Toolkit;

import javax.swing.JFrame;
import javax.swing.JPanel;

import java.awt.BorderLayout;

import javax.swing.JFileChooser;
import javax.swing.JLabel;
import javax.swing.JMenuBar;
import javax.swing.JMenu;
import javax.swing.JMenuItem;
import javax.swing.JButton;
import javax.swing.AbstractAction;
import javax.swing.SwingWorker;

import java.awt.event.ActionEvent;

import javax.swing.Action;

import java.awt.event.ActionListener;
import java.io.File;
import java.util.Random;

 
import lms.tools.ExcelCompile;

import javax.swing.JTextArea;

import java.awt.GridLayout;

import javax.swing.SwingConstants;

import java.awt.FlowLayout;
import java.beans.PropertyChangeEvent;
import java.beans.PropertyChangeListener;

import javax.swing.BoxLayout;
import javax.swing.JProgressBar;

public class CompareWindow {

	private JFrame frame;

    private File excelFile;

	private JButton btnopenButton;

	private JProgressBar progressBar;

	private JButton btnOpenResultButton;

	private JTextArea taskOutput;

	private Task task;
	/**
	 * Launch the application.
	 */
	public static void main(String[] args) {
		EventQueue.invokeLater(new Runnable() {
			public void run() {
				try {
					CompareWindow window = new CompareWindow();
					window.frame.setVisible(true);
				} catch (Exception e) {
					e.printStackTrace();
				}
			}
		});
	}

	/**
	 * Create the application.
	 */
	public CompareWindow() {
		initialize();
	}

	/**
	 * Initialize the contents of the frame.
	 */
	private void initialize() {
		frame = new JFrame();
		frame.setTitle("\u914D\u7F6E\u6BD4\u8F83\u795E\u5668");
		frame.setBounds(100, 100, 450, 300);
		frame.setDefaultCloseOperation(JFrame.EXIT_ON_CLOSE);
		
		JPanel panel = new JPanel();
		frame.getContentPane().add(panel, BorderLayout.CENTER);
		panel.setLayout(new BorderLayout(0, 0));
		
		taskOutput = new JTextArea();
		panel.add(taskOutput);
		
		JPanel panel_1 = new JPanel();
		panel.add(panel_1, BorderLayout.NORTH);
		
		btnopenButton = new JButton("open");
		panel_1.add(btnopenButton);
		btnopenButton.setHorizontalAlignment(SwingConstants.LEFT);
		
		progressBar = new JProgressBar();
		panel_1.add(progressBar);
		
		
		btnOpenResultButton = new JButton("exit");
		panel_1.add(btnOpenResultButton);
		btnOpenResultButton.addActionListener(new ActionListener(){

			@Override
			public void actionPerformed(ActionEvent e) {
				// TODO Auto-generated method stub
//				System.exit(0);
				
				btnOpenResultButton.setEnabled(false);
//			        setCursor(Cursor.getPredefinedCursor(Cursor.WAIT_CURSOR));
			        //Instances of javax.swing.SwingWorker are not reusuable, so
			        //we create new instances as needed.
			        task = new Task();
			        task.addPropertyChangeListener(new PropertyChangeListener(){
				    	@Override
				    	public void propertyChange(PropertyChangeEvent evt) {
				    		// TODO Auto-generated method stub
				    		   if ("progress" == evt.getPropertyName()) {
				    	            int progress = (Integer) evt.getNewValue();
				    	            progressBar.setValue(progress);
				    	            taskOutput.append(String.format(
				    	                    "Completed %d%% of task.\n", task.getProgress()));
				    	        }
				    	}
				    });
			        task.execute();
			}
			
		});
		
		btnOpenResultButton.setHorizontalAlignment(SwingConstants.LEADING);
		btnopenButton.addActionListener(new ActionListener(){

			@Override
			public void actionPerformed(ActionEvent e) {
				// TODO Auto-generated method stub
				openfile();
				// TODO Auto-generated method stub
//				System.exit(0);
				
				btnOpenResultButton.setEnabled(false);
//			        setCursor(Cursor.getPredefinedCursor(Cursor.WAIT_CURSOR));
			        //Instances of javax.swing.SwingWorker are not reusuable, so
			        //we create new instances as needed.
			        task = new Task();
			        task.addPropertyChangeListener(new PropertyChangeListener(){
				    	@Override
				    	public void propertyChange(PropertyChangeEvent evt) {
				    		// TODO Auto-generated method stub
				    		   if ("progress" == evt.getPropertyName()) {
				    	            int progress = (Integer) evt.getNewValue();
				    	            progressBar.setValue(progress);
				    	            taskOutput.append(String.format(
				    	                    "Completed %d%% of task.\n", task.getProgress()));
				    	        }
				    	}
				    });
			        task.execute();
			
			}
			
		});
		
		JMenuBar menuBar = new JMenuBar();
		frame.setJMenuBar(menuBar);
		
		JMenu mnFile = new JMenu("File");
		menuBar.add(mnFile);
		
		JMenuItem mntmOpen = new JMenuItem("Open");
			mnFile.add(mntmOpen);
			mntmOpen.addActionListener(new ActionListener(){

				@Override
				public void actionPerformed(ActionEvent arg0) {
					// TODO Auto-generated method stub
					
					openfile();
					
					
					
					
					
					
					taskOutput.append("\n你的文件已经输出到 "+ ExcelCompile.filename);
					taskOutput.invalidate();
					taskOutput.repaint();
					
				}
				
			});
		
		JMenuItem mntmExit = new JMenuItem("Exit");
		mntmExit.addActionListener(new ActionListener() {
			public void actionPerformed(ActionEvent arg0) {
				System.exit(0);
			}
		});
		mnFile.add(mntmExit);
		
		JMenu mnAbout = new JMenu("Help");
		menuBar.add(mnAbout);
		
		JMenuItem mntmAbout = new JMenuItem("About");
		mnAbout.add(mntmAbout);
	}
	
	void openfile(){
		JFileChooser jfc=new JFileChooser();
		jfc.setFileSelectionMode(JFileChooser.FILES_AND_DIRECTORIES );
		jfc.setCurrentDirectory(new File("d:/"));
		jfc.showDialog(new JLabel(), "选择");
	
		File file=jfc.getSelectedFile();
		if(file.isDirectory()){
			 
		}else if(file.isFile()){
			ExcelCompile  ec= new ExcelCompile();
			ExcelCompile.workpath=file.getParent();
		 
			ec.readFromXLS2003(file);
			ec.formatOut();
			
		}
		
	}

	class Task extends SwingWorker<Void, Void> {
	    /*
	     * Main task. Executed in background thread.
	     */
	    @Override
	    public Void doInBackground() {
	        Random random = new Random();
	        int progress = 0;
	        //Initialize progress property.
	        setProgress(0);
	        while (progress < 100) {
	            //Sleep for up to one second.
	            try {
	                Thread.sleep(random.nextInt(1000));
	            } catch (InterruptedException ignore) {}
	            //Make random progress.
	            progress += random.nextInt(10);
	            setProgress(Math.min(progress, 100));
	        }
	        return null;
	    }

	    /*
	     * Executed in event dispatching thread
	     */
	    @Override
	    public void done() {
	        Toolkit.getDefaultToolkit().beep();
	        btnOpenResultButton.setEnabled(true);
//	        setCursor(null); //turn off the wait cursor
	        taskOutput.append("Done!\n");
	    }
	}

	
}
 
