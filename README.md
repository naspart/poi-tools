# poi-tools
excel导出
```java
@Data
@AllArgsConstructor
@NoArgsConstructor
@ExcelTarget("运单报表")
public class AwbDto {
    /**
     * 运单号
     */
    @ExcelField(name = "运单号", width = 20)
    private String awbNumber;

    /**
     * 代理人
     */
    @ExcelField(name = "代理人", width = 10)
    private String agent;

    /*
     * 用户姓名
     */
    @ExcelField(name = "用户名", width = 10)
    private String userName;
    /*
     * 性别
     */
    @ExcelField(name = "性别", width = 8, replace = {"0_男", "1_女"}, horizontalAlignment = HorizontalAlignment.CENTER)
    private Integer gender;

    /*
     * 身份证号
     */
    @ExcelField(name = "身份证号", width = 20)
    private String idCard;
    /*
     * 工资
     */
    @ExcelField(name = "工资", width = 20, format = "#,##0.00", horizontalAlignment = HorizontalAlignment.GENERAL)
    private double salary;

    /*
     * 工资
     */
    @ExcelField(name = "税", width = 20, format = "#,##0.00", horizontalAlignment = HorizontalAlignment.GENERAL)
    private float tax;

    /*
     * 百分比
     */
    @ExcelField(name = "百分比", width = 10, format = "0.00%", horizontalAlignment = HorizontalAlignment.GENERAL)
    private double percentage;
    /*
     * 生日
     */
    @ExcelField(name = "生日", width = 20, format = "yyyy-MM-dd")
    private LocalDate birthday;
    /*
     * 记录日期
     */
    @ExcelField(name = "记录日期", width = 25, format = "yyyy-MM-dd HH:mm:ss")
    private Calendar recDateTime;
    /*
     * 记录时间
     */

    @ExcelField(name = "记录时间", width = 10, format = "HH:mm:ss")
    private Date recTime;
}
```
```java
import com.naspart.demo_collections.excel.AwbDto;
import com.naspart.poi.ExcelExportBuilder;
import org.springframework.web.bind.annotation.RequestMapping;
import org.springframework.web.bind.annotation.RequestMethod;
import org.springframework.web.bind.annotation.RestController;

import javax.servlet.http.HttpServletResponse;
import java.io.IOException;
import java.io.OutputStream;
import java.time.LocalDate;
import java.time.LocalDateTime;
import java.time.format.DateTimeFormatter;
import java.util.ArrayList;
import java.util.Calendar;
import java.util.Date;
import java.util.List;

@RestController
@RequestMapping(value = "/excel")
public class ExcelController {
    @RequestMapping(value = "/export/v1", produces = {"application/json", "application/octet-stream"}, method = RequestMethod.GET)
    public void export(HttpServletResponse response) throws IOException {
        response.setContentType("octets/stream");
        response.addHeader("Content-Disposition",
                "attachment;filename=" + new String(
                        ("运单月报_" + LocalDateTime.now().format(DateTimeFormatter.ofPattern("yyyyMMddHHmmss"))).getBytes("gb2312"),
                        "ISO8859-1") + ".xlsx");

        OutputStream out = response.getOutputStream();

        ExcelExportBuilder excelExportService = new ExcelExportBuilder();
        excelExportService.build("运单月报(1月)", getData1(), AwbDto.class);
        excelExportService.build("运单月报(2月)", getData1(), AwbDto.class);
        excelExportService.build("运单月报(3月)", getData1(), AwbDto.class);
        excelExportService.build("运单月报(4月)", new Date(), new ExcelDataFunction<Date, AwbDto>() {
            @Override
            public Collection<AwbDto> getData(Date date, int i, int i1) {
                return new ArrayList<>();
            }

            @Override
            public AwbDto convert(AwbDto awbDto) {
                return awbDto;
            }
        }, AwbDto.class);

        try {
            excelExportService.out(out);
        } catch (IOException e) {
            e.printStackTrace();
        }

        try {
            excelExportService.out(out);
        } catch (IOException e) {
            e.printStackTrace();
        }

        out.close();
        System.out.print("Excel导出成功！");
    }

    private List<AwbDto> getData1() {
        List<AwbDto> data = new ArrayList<>();
        data.add(
                new AwbDto(
                        "112",
                        "张三",
                        "李刚",
                        null,
                        "42186846846846",
                        8723456.56,
                        34.5f,
                        0.7895,
                        null,
                        Calendar.getInstance(),
                        new Date())
        );
        data.add(
                new AwbDto(
                        "112",
                        "张三",
                        "李刚",
                        3,
                        "42186846846846",
                        8723456.56,
                        34.5f,
                        0.7895,
                        null,
                        Calendar.getInstance(),
                        new Date())
        );

        for (int i = 0; i < 10000; i++) {
            data.add(
                    new AwbDto(
                            "112-",
                            "张三",
                            "李刚",
                            0,
                            "42186846846846",
                            8723456.56,
                            34.5f,
                            0.7895,
                            LocalDate.now(),
                            Calendar.getInstance(),
                            new Date())
            );
        }
        return data;
    }
}

```