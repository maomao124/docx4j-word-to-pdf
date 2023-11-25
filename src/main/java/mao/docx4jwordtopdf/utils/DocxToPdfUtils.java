package mao.docx4jwordtopdf.utils;

import org.apache.commons.compress.utils.IOUtils;
import org.docx4j.Docx4J;
import org.docx4j.fonts.IdentityPlusMapper;
import org.docx4j.fonts.Mapper;
import org.docx4j.fonts.PhysicalFont;
import org.docx4j.fonts.PhysicalFonts;
import org.docx4j.openpackaging.packages.WordprocessingMLPackage;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

import javax.servlet.http.HttpServletResponse;
import java.io.File;
import java.io.FileOutputStream;
import java.util.UUID;

/**
 * Project name(项目名称)：docx4j-word-to-pdf
 * Package(包名): mao.docx4jwordtopdf.utils
 * Class(类名): DocxToPdfUtils
 * Author(作者）: mao
 * Author QQ：1296193245
 * GitHub：https://github.com/maomao124/
 * Date(创建日期)： 2023/11/25
 * Time(创建时间)： 18:06
 * Version(版本): 1.0
 * Description(描述)： docx转pdf
 */

public class DocxToPdfUtils
{
    private static final Logger log = LoggerFactory.getLogger(DocxToPdfUtils.class);


    /**
     * docx转pdf
     *
     * @param docxPath docx文件路径
     * @param pdfPath  pdf文件路径
     * @throws Exception 异常
     */
    public static void convertDocxToPdf(String docxPath, String pdfPath) throws Exception
    {

        FileOutputStream fileOutputStream = null;
        try
        {
            File file = new File(docxPath);
            fileOutputStream = new FileOutputStream(new File(pdfPath));
            WordprocessingMLPackage mlPackage = WordprocessingMLPackage.load(file);
            setFontMapper(mlPackage);
            Docx4J.toPDF(mlPackage, new FileOutputStream(new File(pdfPath)));
        }
        catch (Exception e)
        {
            e.printStackTrace();
            log.error("docx文档转换为PDF失败");
        }
        finally
        {
            IOUtils.closeQuietly(fileOutputStream);
        }
    }

    /**
     * 加载字体文件（解决linux环境下无中文字体问题）
     *
     * @param mlPackage {@link WordprocessingMLPackage}
     * @throws Exception 异常
     */
    private static void setFontMapper(WordprocessingMLPackage mlPackage) throws Exception
    {
        Mapper fontMapper = new IdentityPlusMapper();
        //加载字体文件（解决linux环境下无中文字体问题）
        if (PhysicalFonts.get("SimSun") == null)
        {
            System.out.println("加载本地SimSun字体库");
            //PhysicalFonts.addPhysicalFonts("SimSun", WordUtils.class.getResource("/fonts/SIMSUN.TTC"));
        }
        fontMapper.put("隶书", PhysicalFonts.get("LiSu"));
        fontMapper.put("宋体", PhysicalFonts.get("SimSun"));
        fontMapper.put("微软雅黑", PhysicalFonts.get("Microsoft Yahei"));
        fontMapper.put("黑体", PhysicalFonts.get("SimHei"));
        fontMapper.put("楷体", PhysicalFonts.get("KaiTi"));
        fontMapper.put("新宋体", PhysicalFonts.get("NSimSun"));
        fontMapper.put("华文行楷", PhysicalFonts.get("STXingkai"));
        fontMapper.put("华文仿宋", PhysicalFonts.get("STFangsong"));
        fontMapper.put("仿宋", PhysicalFonts.get("FangSong"));
        fontMapper.put("幼圆", PhysicalFonts.get("YouYuan"));
        fontMapper.put("华文宋体", PhysicalFonts.get("STSong"));
        fontMapper.put("华文中宋", PhysicalFonts.get("STZhongsong"));
        fontMapper.put("等线", PhysicalFonts.get("SimSun"));
        fontMapper.put("等线 Light", PhysicalFonts.get("SimSun"));
        fontMapper.put("华文琥珀", PhysicalFonts.get("STHupo"));
        fontMapper.put("华文隶书", PhysicalFonts.get("STLiti"));
        fontMapper.put("华文新魏", PhysicalFonts.get("STXinwei"));
        fontMapper.put("华文彩云", PhysicalFonts.get("STCaiyun"));
        fontMapper.put("方正姚体", PhysicalFonts.get("FZYaoti"));
        fontMapper.put("方正舒体", PhysicalFonts.get("FZShuTi"));
        fontMapper.put("华文细黑", PhysicalFonts.get("STXihei"));
        fontMapper.put("宋体扩展", PhysicalFonts.get("simsun-extB"));
        fontMapper.put("仿宋_GB2312", PhysicalFonts.get("FangSong_GB2312"));
        fontMapper.put("新細明體", PhysicalFonts.get("SimSun"));
        //解决宋体（正文）和宋体（标题）的乱码问题
        PhysicalFonts.put("PMingLiU", PhysicalFonts.get("SimSun"));
        PhysicalFonts.put("新細明體", PhysicalFonts.get("SimSun"));
        //宋体&新宋体
        PhysicalFont simsunFont = PhysicalFonts.get("SimSun");
        fontMapper.put("SimSun", simsunFont);
        //设置字体
        mlPackage.setFontMapper(fontMapper);
    }
}
