package com.coreoz.ppt;

import java.awt.image.BufferedImage;
import java.io.ByteArrayOutputStream;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;

import javax.imageio.ImageIO;

public class TestPPTTemplates {

	public void test(String[] args) {

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
