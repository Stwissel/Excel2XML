/** ========================================================================= *
 * Copyright (C)  2017, 2018 Salesforce Inc ( http://www.salesforce.com/      *
 *                            All rights reserved.                            *
 *                                                                            *
 *  @author     Stephan H. Wissel (stw) <swissel@salesforce.com>              *
 *                                       @notessensei                         *
 * @version     1.0                                                           *
 * ========================================================================== *
 *                                                                            *
 * Licensed under the  Apache License, Version 2.0  (the "License").  You may *
 * not use this file except in compliance with the License.  You may obtain a *
 * copy of the License at <http://www.apache.org/licenses/LICENSE-2.0>.       *
 *                                                                            *
 * Unless  required  by applicable  law or  agreed  to  in writing,  software *
 * distributed under the License is distributed on an  "AS IS" BASIS, WITHOUT *
 * WARRANTIES OR  CONDITIONS OF ANY KIND, either express or implied.  See the *
 * License for the  specific language  governing permissions  and limitations *
 * under the License.                                                         *
 *                                                                            *
 * ========================================================================== *
 */
package net.wissel.tools.excel;

import java.io.ByteArrayInputStream;
import java.io.ByteArrayOutputStream;
import java.io.File;
import java.io.IOException;
import java.io.OutputStream;

import javax.xml.parsers.DocumentBuilder;
import javax.xml.parsers.DocumentBuilderFactory;
import javax.xml.parsers.ParserConfigurationException;
import javax.xml.transform.Transformer;
import javax.xml.transform.TransformerException;
import javax.xml.transform.TransformerFactory;
import javax.xml.transform.dom.DOMSource;
import javax.xml.transform.stream.StreamResult;
import javax.xml.transform.stream.StreamSource;

import org.w3c.dom.Document;
import org.xml.sax.SAXException;

/**
 * @author swissel
 *
 */
public class TransformingOutputStream extends OutputStream {

    private final OutputStream finalStream;
    private final OutputStream innerStream;
    private final String       templateName;
    private final boolean      transform;

    public TransformingOutputStream(OutputStream finalStream, String templateName) {
        this.transform = (templateName != null);
        this.templateName = templateName;
        this.finalStream = finalStream;
        this.innerStream = (transform) ? new ByteArrayOutputStream() : finalStream;
    }

    /**
     * @see java.io.OutputStream#write(byte[])
     */
    @Override
    public void write(byte[] b) throws IOException {
        this.innerStream.write(b);
    }

    /**
     * @see java.io.OutputStream#write(byte[], int, int)
     */
    @Override
    public void write(byte[] b, int off, int len) throws IOException {
        this.innerStream.write(b, off, len);
    }

    /**
     * @see java.io.OutputStream#flush()
     */
    @Override
    public void flush() throws IOException {
        this.innerStream.flush();
    }

    /**
     * @see java.io.OutputStream#close()
     */
    @Override
    public void close() throws IOException {

        if (transform) {
            this.executeTransformation();
        }
        finalStream.close();
    }

    private void executeTransformation() {
        try {
            DocumentBuilderFactory factory = DocumentBuilderFactory.newInstance();

            File xsl = new File(this.templateName);
            ByteArrayInputStream xml = new ByteArrayInputStream(
                    ((ByteArrayOutputStream) this.innerStream).toByteArray());

            DocumentBuilder builder = factory.newDocumentBuilder();

            Document document = builder.parse(xml);

            // Use a Transformer for output
            TransformerFactory transformerFactory = TransformerFactory.newInstance();
            StreamSource style = new StreamSource(xsl);
            Transformer transformer = transformerFactory.newTransformer(style);

            DOMSource source = new DOMSource(document);
            StreamResult result = new StreamResult(this.finalStream);
            transformer.transform(source, result);
            this.innerStream.close();
        } catch (ParserConfigurationException | TransformerException | IOException | SAXException e) {
            e.printStackTrace();
        }

    }

    /**
     * @see java.io.OutputStream#write(int)
     */
    @Override
    public void write(int b) throws IOException {
        this.innerStream.write(b);

    }

}
