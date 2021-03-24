package org.apache.poi.ooxml;

import org.apache.poi.hslf.usermodel.HSLFSlideShow;
import org.apache.poi.hslf.usermodel.HSLFSlideShowImpl;
import org.apache.poi.sl.draw.Drawable;
import org.apache.poi.sl.usermodel.Slide;

import java.awt.*;
import java.awt.image.BufferedImage;
import java.io.File;
import java.io.FileInputStream;
import java.io.InputStream;
import java.lang.ref.WeakReference;
import java.util.Date;
import java.util.concurrent.atomic.AtomicReference;

public class TestBisect {
    public static void main(String[] args) throws Exception {
        String fileName = "alfresco.vgregion.se_alfresco_service_vgr_storage_node_content_workspace_spacesstore_a1dc0dc0-b6f6-4890-8dc6-e4dd029764f1_tom_20fl_c3_b6desschema_20nytt.ppt_a=false&guest=true&native=true.ppt";
        String ROOT_DIR = "/opt/CommonCrawl/download2";
        File file = new File(ROOT_DIR, fileName);

        System.out.println("Handling file: " + file);

        if (!file.exists()) {
            throw new IllegalStateException("File not found: " + file.getAbsolutePath());
        }

        final AtomicReference<Exception> exc = new AtomicReference<>();

        Thread thread = new Thread(() -> {
            //System.setProperty("org.apache.poi.util.POILogger", "org.apache.poi.util.SystemOutLogger");
            try (InputStream stream = new FileInputStream(file)) {
                HSLFSlideShowImpl slide = new HSLFSlideShowImpl(stream);
                HSLFSlideShow ss = new HSLFSlideShow(slide);
                Dimension pgSize = ss.getPageSize();

                for (Slide<?, ?> s : ss.getSlides()) {
                    BufferedImage img = new BufferedImage(pgSize.width, pgSize.height, BufferedImage.TYPE_INT_ARGB);
                    Graphics2D graphics = img.createGraphics();

                    // default rendering options
                    graphics.setRenderingHint(RenderingHints.KEY_ANTIALIASING, RenderingHints.VALUE_ANTIALIAS_ON);
                    graphics.setRenderingHint(RenderingHints.KEY_RENDERING, RenderingHints.VALUE_RENDER_QUALITY);
                    graphics.setRenderingHint(RenderingHints.KEY_INTERPOLATION, RenderingHints.VALUE_INTERPOLATION_BICUBIC);
                    graphics.setRenderingHint(RenderingHints.KEY_FRACTIONALMETRICS, RenderingHints.VALUE_FRACTIONALMETRICS_ON);
                    graphics.setRenderingHint(Drawable.BUFFERED_IMAGE, new WeakReference<>(img));

                    // draw stuff
                    s.draw(graphics);

                    graphics.dispose();
                    img.flush();
                }
            } catch (Exception e) {
                exc.set(e);
            }

        });

        thread.setDaemon(true);
        thread.start();
        for (int i = 0; i < 30; i++) {
            System.out.println("Waiting for thread... " + new Date());
            thread.join(10_000);

            if (!thread.isAlive()) {
                break;
            }
        }

        if (exc.get() != null) {
            throw exc.get();
        }

        if (thread.isAlive()) {
            throw new IllegalStateException("Thread did not finish in time");
        }
    }
}
