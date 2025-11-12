import java.awt.geom.Path2D;
import java.awt.geom.Point2D;
import java.awt.geom.Rectangle2D;
import java.io.File;
import java.io.IOException;
import java.io.Writer;
import java.nio.charset.StandardCharsets;
import java.util.ArrayDeque;
import java.util.ArrayList;
import java.util.Arrays;
import java.util.Collection;
import java.util.Deque;
import java.util.HashMap;
import java.util.LinkedHashMap;
import java.util.List;
import java.util.Locale;
import java.util.Map;
import java.util.Set;
import java.util.Collections;
import java.util.TreeSet;

import org.apache.pdfbox.Loader;
import org.apache.pdfbox.contentstream.PDFGraphicsStreamEngine;
import org.apache.pdfbox.contentstream.operator.OperatorName;
import org.apache.pdfbox.cos.COSName;
import org.apache.pdfbox.pdmodel.PDDocument;
import org.apache.pdfbox.pdmodel.PDPage;
import org.apache.pdfbox.pdmodel.graphics.color.PDColor;
import org.apache.pdfbox.pdmodel.graphics.image.PDImage;
import org.apache.pdfbox.text.PDFTextStripper;
import org.apache.pdfbox.text.TextPosition;

public final class GeometryExtractor {
    public static void main(String[] args) throws Exception {
        Map<String, String> options = parseArgs(args);
        String pdfPath = require(options, "--pdf");
        String outputPath = require(options, "--output");
        String pagesOption = options.get("--pages");

        Set<Integer> pagesToProcess = parsePages(pagesOption);

        try (PDDocument document = Loader.loadPDF(new File(pdfPath))) {
            if (pagesToProcess.isEmpty()) {
                for (int i = 1; i <= document.getNumberOfPages(); i++) {
                    pagesToProcess.add(i);
                }
            }

            WordCollector wordCollector = new WordCollector();
            wordCollector.setSortByPosition(true);
            wordCollector.collect(document);

            List<PagePayload> pages = new ArrayList<>();
            for (int pageNumber : pagesToProcess) {
                if (pageNumber < 1 || pageNumber > document.getNumberOfPages()) {
                    continue;
                }

                PDPage page = document.getPage(pageNumber - 1);
                GraphicsCollector graphicsCollector = new GraphicsCollector(page);
                graphicsCollector.processPage(page);

                pages.add(new PagePayload(
                    pageNumber,
                    wordCollector.getWords(pageNumber),
                    graphicsCollector.getGraphics()));
            }

            writeJson(outputPath, pages);
        }
    }

    private static Map<String, String> parseArgs(String[] args) {
        Map<String, String> map = new HashMap<>();
        for (int i = 0; i < args.length; i++) {
            if (args[i].startsWith("--")) {
                String key = args[i];
                String value = (i + 1) < args.length ? args[i + 1] : "";
                if (!value.startsWith("--")) {
                    map.put(key, value);
                    i++;
                } else {
                    map.put(key, "");
                }
            }
        }
        return map;
    }

    private static String require(Map<String, String> options, String key) {
        String value = options.get(key);
        if (value == null || value.isEmpty()) {
            throw new IllegalArgumentException("Missing option " + key);
        }
        return value;
    }

    private static Set<Integer> parsePages(String input) {
        Set<Integer> pages = new TreeSet<>();
        if (input == null || input.isEmpty()) {
            return pages;
        }

        for (String token : input.split(",")) {
            token = token.trim();
            if (token.isEmpty()) continue;
            if (token.contains("-")) {
                String[] bounds = token.split("-");
                int start = Integer.parseInt(bounds[0].trim());
                int end = Integer.parseInt(bounds[1].trim());
                if (end < start) {
                    int tmp = start;
                    start = end;
                    end = tmp;
                }
                for (int i = start; i <= end; i++) {
                    pages.add(i);
                }
            } else {
                pages.add(Integer.parseInt(token));
            }
        }

        return pages;
    }

    private static void writeJson(String path, List<PagePayload> pages) throws IOException {
        try (Writer writer = new java.io.OutputStreamWriter(new java.io.FileOutputStream(path), StandardCharsets.UTF_8)) {
            writer.write("{\"pages\":[");
            for (int i = 0; i < pages.size(); i++) {
                if (i > 0) writer.write(",");
                pages.get(i).writeJson(writer);
            }
            writer.write("]}");
        }
    }

    private static final class PagePayload {
        private final int page;
        private final List<WordInfo> words;
        private final List<GraphicInfo> graphics;

        private PagePayload(int page, List<WordInfo> words, List<GraphicInfo> graphics) {
            this.page = page;
            this.words = words;
            this.graphics = graphics;
        }

        void writeJson(Writer writer) throws IOException {
            writer.write("{\"page\":");
            writer.write(Integer.toString(page));
            writer.write(",\"words\":[");
            for (int i = 0; i < words.size(); i++) {
                if (i > 0) writer.write(",");
                words.get(i).writeJson(writer);
            }
            writer.write("],\"graphics\":[");
            for (int i = 0; i < graphics.size(); i++) {
                if (i > 0) writer.write(",");
                graphics.get(i).writeJson(writer);
            }
            writer.write("]}");
        }
    }

    private static final class WordCollector extends PDFTextStripper {
        private final Map<Integer, List<WordInfo>> wordsByPage = new LinkedHashMap<>();
        private WordBuilder builder;
        private TextPosition previous;

        WordCollector() throws IOException {
            super();
        }

        void collect(PDDocument document) throws IOException {
            setStartPage(1);
            setEndPage(document.getNumberOfPages());
            getText(document);
        }

        List<WordInfo> getWords(int page) {
            List<WordInfo> list = wordsByPage.get(page);
            return list == null ? Collections.<WordInfo>emptyList() : list;
        }

        @Override
        protected void startPage(PDPage page) throws IOException {
            super.startPage(page);
            wordsByPage.computeIfAbsent(getCurrentPageNo(), key -> new ArrayList<>());
            builder = null;
            previous = null;
        }

        @Override
        protected void writeString(String text, List<TextPosition> textPositions) throws IOException {
            for (TextPosition position : textPositions) {
                String unicode = position.getUnicode();
                if (unicode == null || unicode.isEmpty()) {
                    continue;
                }

                char ch = unicode.charAt(0);
                if (Character.isWhitespace(ch)) {
                    flushWord();
                    previous = null;
                    continue;
                }

                if (builder == null || needsBreak(position)) {
                    flushWord();
                    builder = new WordBuilder(position);
                } else {
                    builder.append(position);
                }
                previous = position;
            }

            flushWord();
        }

        private boolean needsBreak(TextPosition current) {
            if (previous == null) return true;
            float prevEnd = previous.getXDirAdj() + previous.getWidthDirAdj();
            float gap = current.getXDirAdj() - prevEnd;
            float threshold = current.getWidthDirAdj() * 0.5f;
            return gap > threshold;
        }

        private void flushWord() {
            if (builder == null || builder.isEmpty()) return;
            int page = getCurrentPageNo();
            wordsByPage.computeIfAbsent(page, key -> new ArrayList<>())
                .add(builder.build());
            builder = null;
        }
    }

    private static final class WordBuilder {
        private final StringBuilder text = new StringBuilder();
        private double minX;
        private double minY;
        private double maxX;
        private double maxY;
        private double fontSizeSum;
        private int count;

        WordBuilder(TextPosition position) {
            append(position);
        }

        void append(TextPosition position) {
            text.append(position.getUnicode());
            double x = position.getXDirAdj();
            double y = position.getYDirAdj();
            double w = position.getWidthDirAdj();
            double h = position.getHeightDir();

            if (count == 0) {
                minX = x;
                minY = y - h;
                maxX = x + w;
                maxY = y;
            } else {
                minX = Math.min(minX, x);
                minY = Math.min(minY, y - h);
                maxX = Math.max(maxX, x + w);
                maxY = Math.max(maxY, y);
            }

            fontSizeSum += position.getFontSizeInPt();
            count++;
        }

        boolean isEmpty() {
            return count == 0 || text.toString().trim().isEmpty();
        }

        WordInfo build() {
            double width = Math.max(0, maxX - minX);
            double height = Math.max(0, maxY - minY);
            double fontSize = count == 0 ? 0 : fontSizeSum / count;
            return new WordInfo(
                text.toString(),
                minX,
                minY,
                width,
                height,
                fontSize);
        }
    }

    private static final class WordInfo {
        private final String text;
        private final double x;
        private final double y;
        private final double width;
        private final double height;
        private final double fontSize;

        WordInfo(String text, double x, double y, double width, double height, double fontSize) {
            this.text = text;
            this.x = x;
            this.y = y;
            this.width = width;
            this.height = height;
            this.fontSize = fontSize;
        }

        void writeJson(Writer writer) throws IOException {
            writer.write("{\"text\":\"");
            writer.write(escape(text));
            writer.write("\",\"x\":");
            writer.write(doubleToString(x));
            writer.write(",\"y\":");
            writer.write(doubleToString(y));
            writer.write(",\"width\":");
            writer.write(doubleToString(width));
            writer.write(",\"height\":");
            writer.write(doubleToString(height));
            writer.write(",\"fontSize\":");
            writer.write(doubleToString(fontSize));
            writer.write("}");
        }
    }

    private static final class GraphicsCollector extends PDFGraphicsStreamEngine {
        private final List<GraphicInfo> graphics = new ArrayList<>();
        private Path2D currentPath;
        private Point2D currentPoint;

        GraphicsCollector(PDPage page) {
            super(page);
        }

        List<GraphicInfo> getGraphics() {
            return graphics;
        }

        @Override
        public void appendRectangle(Point2D p0, Point2D p1, Point2D p2, Point2D p3) {
            ensurePath();
            currentPath.moveTo(p0.getX(), p0.getY());
            currentPath.lineTo(p1.getX(), p1.getY());
            currentPath.lineTo(p2.getX(), p2.getY());
            currentPath.lineTo(p3.getX(), p3.getY());
            currentPath.closePath();
            currentPoint = p0;
        }

        @Override
        public void drawImage(PDImage pdImage) throws IOException {
            // ignore raster images
        }

        @Override
        public void clip(int windingRule) {
            // ignore clipping
        }

        @Override
        public void moveTo(float x, float y) throws IOException {
            ensurePath();
            currentPath.moveTo(x, y);
            currentPoint = new Point2D.Double(x, y);
        }

        @Override
        public void lineTo(float x, float y) throws IOException {
            ensurePath();
            currentPath.lineTo(x, y);
            currentPoint = new Point2D.Double(x, y);
        }

        @Override
        public void curveTo(float x1, float y1, float x2, float y2, float x3, float y3) throws IOException {
            ensurePath();
            currentPath.curveTo(x1, y1, x2, y2, x3, y3);
            currentPoint = new Point2D.Double(x3, y3);
        }

        @Override
        public Point2D getCurrentPoint() throws IOException {
            return currentPoint;
        }

        @Override
        public void closePath() throws IOException {
            if (currentPath != null) {
                currentPath.closePath();
            }
        }

        @Override
        public void endPath() throws IOException {
            currentPath = null;
            currentPoint = null;
        }

        @Override
        public void strokePath() throws IOException {
            registerPath("stroke");
        }

        @Override
        public void fillPath(int windingRule) throws IOException {
            registerPath("fill");
        }

        @Override
        public void fillAndStrokePath(int windingRule) throws IOException {
            registerPath("fill-stroke");
        }

        @Override
        public void shadingFill(COSName shadingName) throws IOException {
            // ignore shading
        }

        private void ensurePath() {
            if (currentPath == null) {
                currentPath = new Path2D.Double();
            }
        }

        private void registerPath(String type) throws IOException {
            if (currentPath == null) return;
            Rectangle2D bounds = currentPath.getBounds2D();
            currentPath = null;
            currentPoint = null;

            if (bounds.isEmpty()) return;

            org.apache.pdfbox.pdmodel.graphics.state.PDGraphicsState state = getGraphicsState();
            double strokeWidth = state.getLineWidth();
            PDColor strokeColor = state.getStrokingColor();
            PDColor nonStrokeColor = state.getNonStrokingColor();

            graphics.add(new GraphicInfo(
                type,
                bounds.getX(),
                bounds.getY(),
                bounds.getWidth(),
                bounds.getHeight(),
                strokeWidth,
                toColor(strokeColor),
                toColor(nonStrokeColor)));
        }
    }

    private static final class GraphicInfo {
        private final String type;
        private final double x;
        private final double y;
        private final double width;
        private final double height;
        private final double strokeWidth;
        private final String strokeColor;
        private final String fillColor;

        GraphicInfo(
            String type,
            double x,
            double y,
            double width,
            double height,
            double strokeWidth,
            String strokeColor,
            String fillColor) {
            this.type = type;
            this.x = x;
            this.y = y;
            this.width = width;
            this.height = height;
            this.strokeWidth = strokeWidth;
            this.strokeColor = strokeColor;
            this.fillColor = fillColor;
        }

        void writeJson(Writer writer) throws IOException {
            writer.write("{\"type\":\"");
            writer.write(type);
            writer.write("\",\"x\":");
            writer.write(doubleToString(x));
            writer.write(",\"y\":");
            writer.write(doubleToString(y));
            writer.write(",\"width\":");
            writer.write(doubleToString(width));
            writer.write(",\"height\":");
            writer.write(doubleToString(height));
            writer.write(",\"strokeWidth\":");
            writer.write(doubleToString(strokeWidth));
            writer.write(",\"strokeColor\":");
            writeNullableString(writer, strokeColor);
            writer.write(",\"fillColor\":");
            writeNullableString(writer, fillColor);
            writer.write("}");
        }
    }

    private static void writeNullableString(Writer writer, String value) throws IOException {
        if (value == null) {
            writer.write("null");
        } else {
            writer.write("\"");
            writer.write(escape(value));
            writer.write("\"");
        }
    }

    private static String doubleToString(double value) {
        if (Double.isNaN(value) || Double.isInfinite(value)) {
            return "0";
        }
        return String.format(Locale.US, "%.6f", value);
    }

    private static String escape(String value) {
        StringBuilder sb = new StringBuilder();
        for (char c : value.toCharArray()) {
            switch (c) {
                case '\\':
                    sb.append("\\\\");
                    break;
                case '"':
                    sb.append("\\\"");
                    break;
                case '\n':
                    sb.append("\\n");
                    break;
                case '\r':
                    sb.append("\\r");
                    break;
                case '\t':
                    sb.append("\\t");
                    break;
                default:
                    if (c < 0x20) {
                        sb.append(String.format("\\u%04x", (int) c));
                    } else {
                        sb.append(c);
                    }
                    break;
            }
        }
        return sb.toString();
    }

    private static String toColor(PDColor color) throws IOException {
        if (color == null) return null;
        if (color.getColorSpace() == null) return null;
        float[] rgb = color.getColorSpace().toRGB(color.getComponents());
        int r = clampColor(rgb, 0);
        int g = clampColor(rgb, 1);
        int b = clampColor(rgb, 2);
        return String.format("#%02x%02x%02x", r, g, b);
    }

    private static int clampColor(float[] rgb, int index) {
        if (rgb == null || index >= rgb.length) return 0;
        int value = Math.round(rgb[index] * 255f);
        return Math.max(0, Math.min(255, value));
    }
}
