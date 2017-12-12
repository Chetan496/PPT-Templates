package com.coreoz.ppt;

import java.awt.Color;
import java.awt.geom.Rectangle2D;
import java.awt.image.BufferedImage;
import java.io.ByteArrayOutputStream;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.ArrayList;
import java.util.Arrays;
import java.util.List;
import java.util.stream.Collectors;

import javax.imageio.ImageIO;

import org.apache.poi.xslf.usermodel.XMLSlideShow;
import org.apache.poi.xslf.usermodel.XSLFShape;
import org.apache.poi.xslf.usermodel.XSLFSlide;
import org.apache.poi.xslf.usermodel.XSLFTextShape;

public class TestPPTTemplates {

	
	public static void main(String[] args) {
		
		testMultipleParaGen();
	}
	
	
	public static void testMultipleParaGen( ) {
		
		List<String> myTextList = new ArrayList<String>( Arrays.asList(  "apples", "nectarines", "guavas" ,
				"guavas" ,
				"guavas" ,
				"guavas" ,
				"guavas" ,
				"guavas" ,
				"guavas" ,
				"guavas" ,
				"guavas" ,
				"guavas" ,
				"guavas" ,
				"guavas" ,
				"guavas" ,
				"guavas" ,
				"guavas" ,
				"guavas" ,
				"guavas" ,
				"guavas" ,
				"guavas" ,
				"guavas" ,
				"guavas" ,
				"guavas" ,
				"guavas" ,
				"guavas" ,
				"guavas" ,
				"guavas" ) )  ;
		
		try {
			final XMLSlideShow xmlSlideShow = new XMLSlideShow(new FileInputStream("E:\\test.pptx"));
			XSLFSlide slide =  xmlSlideShow.getSlides().get(0) ;
			
			List<XSLFShape> shapeList = slide.getShapes()
					                          .stream()
					                          .filter( (shape) -> shape instanceof XSLFTextShape  )
					                          .collect(Collectors.toCollection(ArrayList::new)) ;
			
			
			java.awt.Dimension pgsize = xmlSlideShow.getPageSize();
		    int pgx = pgsize.width; //slide width in points
		    int pgy = pgsize.height; //slide height in points
			
		    
		    //you can detect when the anchor of the containing text box is overflowing.
		    //in that case, create new slide and create the shape on new slide.
			if(shapeList.get(0) == null) {
				return;
			}else {
			
				CreateMultipleTextParagraphs.createMultipleTextParagraphs(myTextList,   (XSLFTextShape) shapeList.get(0) );
				
			}
			
			xmlSlideShow.write(new FileOutputStream("E:\\test.pptx"));
			xmlSlideShow.close();
			
			
			
			
		} catch (FileNotFoundException e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		} catch (IOException e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		}
		
		
		

	}
	
	
	
	public static void testPPTGenFromTemplate(String[] args) {

		final String imagePath = "C:\\Users\\Chetan\\Pictures\\notfound.jpg";
		byte[] jpgByteArray = null;

		try {
			BufferedImage image = ImageIO.read(new File(imagePath));
			ByteArrayOutputStream b = new ByteArrayOutputStream();
			ImageIO.write(image, "jpg", b);

			jpgByteArray = b.toByteArray();
			
			
			// byte[] imageData - needs to be found.

			try (FileOutputStream out = new FileOutputStream("E:\\generated.pptx")) {
				try {
					new PptMapper().text("title", "Hello").text("secondline", "SecondLine")
									.text("firstline", "FirstLine")
									.image("image1", jpgByteArray, PptImageReplacementMode.RESIZE_ONLY )
									.styleText( "sampleLink", textRun -> {
										textRun.setBold(true);
										textRun.setFontColor(Color.RED);
									})
									.styleText(  "sampleLink2", textRun -> {
										textRun.setBold(true);
										textRun.setItalic(true);
										textRun.setFontColor(Color.GREEN);
									}) /*this is how you make use of args. */
									.styleShape("sampleShape",  (arg, shape) -> {
										Rectangle2D shapeAnchor = shape.getAnchor();
										
										shape.setAnchor(new Rectangle2D.Double(
											shapeAnchor.getX(),
											shapeAnchor.getY(),
											
											shapeAnchor.getWidth() * 1.5,
											shapeAnchor.getHeight()
										));
										
										if("green".equals(arg)) {
											shape.setFillColor(Color.GREEN);	
										}else {
											shape.setFillColor(Color.BLACK);
										}
										
									}) /*this is how you make use of args */
									.text("varwitharg",   argument -> {
										if("makemegreen".equals(argument)) {
											return "green text";
										}
										
										return "non green";
									})
							        .processTemplate(new FileInputStream("E:\\Template.pptx")).write(out);
					
					
					
				} catch (FileNotFoundException e) {
					// TODO Auto-generated catch block
					e.printStackTrace();
				} catch (IOException e) {
					// TODO Auto-generated catch block
					e.printStackTrace();
				}
			} catch (FileNotFoundException e1) {
				// TODO Auto-generated catch block
				e1.printStackTrace();
			} catch (IOException e1) {
				// TODO Auto-generated catch block
				e1.printStackTrace();
			}
			

		} catch (IOException e2) {
			// TODO Auto-generated catch block
			e2.printStackTrace();
		}



	}
	
	
	

}
