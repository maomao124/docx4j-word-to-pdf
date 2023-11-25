package mao.docx4jwordtopdf.utils;

import org.junit.jupiter.api.Test;

import static org.junit.jupiter.api.Assertions.*;

/**
 * Project name(项目名称)：docx4j-word-to-pdf
 * Package(包名): mao.docx4jwordtopdf.utils
 * Class(测试类名): DocxToPdfUtilsTest
 * Author(作者）: mao
 * Author QQ：1296193245
 * GitHub：https://github.com/maomao124/
 * Date(创建日期)： 2023/11/25
 * Time(创建时间)： 18:14
 * Version(版本): 1.0
 * Description(描述)： 测试类
 */

class DocxToPdfUtilsTest
{

    @Test
    void convertDocxToPdf() throws Exception
    {
        DocxToPdfUtils.convertDocxToPdf("./test.docx","./test.pdf");
    }

    @Test
    void convertDocxToPdf2() throws Exception
    {
        //相对复杂的docx
        DocxToPdfUtils.convertDocxToPdf("./out.docx","./out.pdf");
    }
}
