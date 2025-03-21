package manager

import model.Project
import model.Member
import org.apache.poi.ss.usermodel.DataFormatter
import org.apache.poi.ss.usermodel.Workbook
import org.apache.poi.ss.usermodel.WorkbookFactory
import org.apache.poi.xssf.usermodel.XSSFWorkbook
import java.io.FileInputStream
import java.io.FileOutputStream
import java.io.IOException
import java.time.LocalDate
import java.time.format.DateTimeFormatter
import java.time.format.DateTimeParseException

object ExcelManager {
    //文件目录
    private const val FILE_PATH="workbook.xlsx"

    //未发送的项目信息
    private var _projectList= mutableListOf<Project>()
    val projectList get() = _projectList

    //已发送的项目信息
    private var _finishedList= mutableListOf<Project>()
    private val finishedList get() = _finishedList

    //用户和钉钉ID的信息
    private var _memberList= mutableListOf<Member>()
    private val memberList get() = _memberList

    //用户和钉钉ID信息的缓存映射（在读取成员时初始化）
    private val _memberMap = mutableMapOf<String, String>()
    val memberMap: Map<String, String> get() = _memberMap


    //初始化表格
    fun init(){
        val filePath = FILE_PATH
        val workbook: Workbook = try {
            // 读取已有文件
            //WorkbookFactory.create(FileInputStream(filePath))
            FileInputStream(filePath).use { WorkbookFactory.create(it) }
        } catch (e: Exception) {
            // 文件不存在时创建新工作簿
            XSSFWorkbook()
        }

        // 获取已有sheet或创建新sheet
        val sheet = workbook.getSheet("未发送") ?: workbook.createSheet("未发送")
        val sheet2 = workbook.getSheet("已发送") ?: workbook.createSheet("已发送")
        val sheet3 = workbook.getSheet("人员信息") ?: workbook.createSheet("人员信息")

        //分别创建表格题头
        var row=sheet.createRow(0)
        for(i in 0..4){
            val cell=row.createCell(i)
            when(i){
                0->cell.setCellValue("项目名称")
                1->cell.setCellValue("发送时间")
                2->cell.setCellValue("发送消息")
                3->cell.setCellValue("是否已发送")
                4->cell.setCellValue("需@的人员姓名")
            }
        }

        row=sheet2.createRow(0)
        for(i in 0..4){
            val cell=row.createCell(i)
            when(i){
                0->cell.setCellValue("项目名称")
                1->cell.setCellValue("发送时间")
                2->cell.setCellValue("发送消息")
                3->cell.setCellValue("是否已发送")
                4->cell.setCellValue("需@的人员姓名")
            }
        }

        row=sheet3.createRow(0)
        for(i in 0..1){
            val cell=row.createCell(i)
            when(i){
                0->cell.setCellValue("用户名")
                1->cell.setCellValue("钉钉ID")
            }
        }

        // 保存文件
        FileOutputStream(filePath).use {
            workbook.write(it)
        }

        //自动调整列长度
        for (i in 0..5){
            sheet.autoSizeColumn(i)
            sheet2.autoSizeColumn(i)
            sheet3.autoSizeColumn(i)
        }

        //关闭工作簿
        try {
            workbook.close()
        }catch (e: IOException){
            println(e)
        }
    }


    //从未发送的表格读取信息
    fun readFromProjectList(){
        val filePath = FILE_PATH
        val workbook: Workbook = try {
            // 读取已有文件
            FileInputStream(filePath).use { WorkbookFactory.create(it) }
        } catch (e: Exception) {
            // 文件不存在时创建新工作簿
            XSSFWorkbook()
        }

        // 获取已有sheet或创建新sheet
        val sheet = workbook.getSheet("未发送") ?: workbook.createSheet("未发送")

        //drop(1)代表跳过首行
        _projectList= (sheet?.drop(1)?.mapNotNull { row ->
            //对单元格内的日期进行转换
            //使用 DataFormatter 获取格式化后的字符串
            val dataFormatter = DataFormatter()
            val timeString = dataFormatter.formatCellValue(row.getCell(1))
            //println(timeString)

            //由于手动输入Excel的日期会被自动转换，故引入多个格式进行解析
            val formatters = listOf(
                DateTimeFormatter.ofPattern("yyyy/M/d"),    // 程序写入的格式
                DateTimeFormatter.ofPattern("M/d/yy"),      // Excel自动转换的短格式
                DateTimeFormatter.ofPattern("yyyy-MM-dd"),  // ISO标准格式
                DateTimeFormatter.ofPattern("dd-MMM-yy")    // Excel其他本地格式（如英文环境）
            )

            var parsedDate: LocalDate? = null
            for (formatter in formatters) {
                try {
                    parsedDate = LocalDate.parse(timeString, formatter)
                    break // 解析成功则跳出循环
                } catch (e: DateTimeParseException) {
                    // 忽略错误，继续尝试下一格式
                }
            }

            //尝试将单元格信息读取进Project数组
            try {
                Project(
                    name = row.getCell(0).toString(), //项目名称
                    time = parsedDate?:LocalDate.now(), // 发送时间
                    message = row.getCell(2).toString(), // 发送消息
                    isMessageSend =row?.getCell(3)?.toString()?.toBoolean()?:false, //做出空判断，否则不填写此内容的单元格将无法被读取
                    memberName = row?.getCell(4)?.toString() ?: "" // 即使为空也返回空字符串
                )
            } catch (e: Exception) {
                null  // 忽略格式错误行
            }
        } ?: emptyList()).toMutableList()

        //关闭工作簿
        try {
            workbook.close()
        }catch (e: IOException){
            println(e)
        }
    }


    //将projectList内的内容覆盖回未发送表格
    fun storeToProjectSheet(){
        val filePath = FILE_PATH
        val workbook: Workbook = try {
            // 读取已有文件
            //WorkbookFactory.create(FileInputStream(filePath))
            FileInputStream(filePath).use { WorkbookFactory.create(it) }
        } catch (e: Exception) {
            // 文件不存在时创建新工作簿
            XSSFWorkbook()
        }

        // 获取已有sheet或创建新sheet
        val sheet = workbook.getSheet("未发送") ?: workbook.createSheet("未发送")

        // 清空原有数据（保留标题行）
        for (i in sheet.lastRowNum downTo 1) { // 从最后一行向上删除，保留第0行
            sheet.removeRow(sheet.getRow(i))
        }

        for (i in 0 until projectList.size){
            val row=sheet.createRow(i+1) //忽略标题行
            for (j in 0..4){
                when(j){
                    0-> { //项目名称列
                        val cell=row.createCell(0)
                        cell.setCellValue(projectList[i].name)
                    }
                    1->{ //发送时间列
                        val cell=row.createCell(1)
                        cell.setCellValue(projectList[i].time.format(DateTimeFormatter.ofPattern("yyyy/M/d"))) //格式化时间
                    }
                    2->{ //发送消息列
                        val cell=row.createCell(2)
                        cell.setCellValue(projectList[i].message)
                    }
                    3->{ //是否已发送列
                        val cell=row.createCell(3)
                        cell.setCellValue(projectList[i].isMessageSend)
                    }
                    4->{ //需@的UserID列
                        val cell=row.createCell(4)
                        cell.setCellValue(projectList[i].memberName?:"") //处理空值
                    }
                }
            }
        }

        // 保存文件
        FileOutputStream(filePath).use {
            workbook.write(it)
        }

        //自动调整列长度，注意要写入以后才有效
        for (i in 0..5){
            sheet.autoSizeColumn(i)
        }

        //关闭工作簿
        try {
            workbook.close()
        }catch (e: IOException){
            println(e)
        }
    }

    //更新数据，把已发送的消息从projectList中剔除，存入finishedList
    fun updateData(){
        _projectList.removeAll { project ->
            if (project.isMessageSend) {
                _finishedList.add(project)
                true // 返回 true 表示需要从 projectList 中移除
            } else {
                false
            }
        }
    }

    //从已发送表格读取信息
    fun readFromFinishedList(){
        val filePath = FILE_PATH
        val workbook: Workbook = try {
            // 读取已有文件
            FileInputStream(filePath).use { WorkbookFactory.create(it) }
        } catch (e: Exception) {
            // 文件不存在时创建新工作簿
            XSSFWorkbook()
        }

        // 获取已有sheet或创建新sheet
        val sheet = workbook.getSheet("已发送") ?: workbook.createSheet("已发送")

        //drop(1)代表跳过首行
        _finishedList= (sheet?.drop(1)?.mapNotNull { row ->
            //对单元格内的日期进行转换
            //使用 DataFormatter 获取格式化后的字符串
            val dataFormatter = DataFormatter()
            val timeString = dataFormatter.formatCellValue(row.getCell(1))
            //println(timeString)

            //由于手动输入Excel的日期会被自动转换，故引入多个格式进行解析
            val formatters = listOf(
                DateTimeFormatter.ofPattern("yyyy/M/d"),    // 程序写入的格式
                DateTimeFormatter.ofPattern("M/d/yy"),      // Excel自动转换的短格式
                DateTimeFormatter.ofPattern("yyyy-MM-dd"),  // ISO标准格式
                DateTimeFormatter.ofPattern("dd-MMM-yy")    // Excel其他本地格式（如英文环境）
            )

            var parsedDate: LocalDate? = null
            for (formatter in formatters) {
                try {
                    parsedDate = LocalDate.parse(timeString, formatter)
                    break // 解析成功则跳出循环
                } catch (e: DateTimeParseException) {
                    // 忽略错误，继续尝试下一格式
                }
            }

            //尝试将单元格信息写入finishedList
            try {
                Project(
                    name = row.getCell(0).toString(), //项目名称
                    time = parsedDate?:LocalDate.now(), // 发送时间
                    message = row.getCell(2).toString(), // 发送消息
                    isMessageSend =row?.getCell(3)?.toString()?.toBoolean()?:false, //做出空判断，否则不填写此内容的单元格将无法被读取
                    memberName = row?.getCell(4)?.toString() //做出空判断，否则不填写此内容的单元格将无法被读取
                )
            } catch (e: Exception) {
                null  // 忽略格式错误行
            }
        } ?: emptyList()).toMutableList()

        //关闭工作簿
        try {
            workbook.close()
        }catch (e: IOException){
            println(e)
        }
    }

    //将finishedList内的内容覆盖回已发送表格
    fun storeToFinishedSheet(){
        val filePath = FILE_PATH
        val workbook: Workbook = try {
            // 读取已有文件
            //WorkbookFactory.create(FileInputStream(filePath))
            FileInputStream(filePath).use { WorkbookFactory.create(it) }
        } catch (e: Exception) {
            // 文件不存在时创建新工作簿
            XSSFWorkbook()
        }

        // 获取已有sheet或创建新sheet
        val sheet = workbook.getSheet("已发送") ?: workbook.createSheet("已发送")

        // 清空原有数据（保留标题行）
        for (i in sheet.lastRowNum downTo 1) { // 从最后一行向上删除，保留第0行
            sheet.removeRow(sheet.getRow(i))
        }

        //println(finishedList)

        for (i in 0 until finishedList.size) {
            val row = sheet.createRow(i + 1) //忽略标题行
            for (j in 0..4) {
                when (j) {
                    0 -> { //项目名称列
                        val cell = row.createCell(0)
                        cell.setCellValue(finishedList[i].name)
                    }

                    1 -> { //发送时间列
                        val cell = row.createCell(1)
                        cell.setCellValue(finishedList[i].time.format(DateTimeFormatter.ofPattern("yyyy/M/d"))) //格式化时间
                    }

                    2 -> { //发送消息列
                        val cell = row.createCell(2)
                        cell.setCellValue(finishedList[i].message)
                    }

                    3 -> { //是否已发送列
                        val cell = row.createCell(3)
                        cell.setCellValue(finishedList[i].isMessageSend)
                    }

                    4 -> { //需@的UserID列
                        val cell = row.createCell(4)
                        cell.setCellValue(finishedList[i].memberName ?: "") //处理空值
                    }
                }
            }
        }
        // 保存文件
        FileOutputStream(filePath).use {
            workbook.write(it)
        }

        //自动调整列长度，注意要写入以后才有效
        for (i in 0..5){
            sheet.autoSizeColumn(i)
        }

        //关闭工作簿
        try {
            workbook.close()
        }catch (e: IOException){
            println(e)
        }
    }


    //读取用户信息表格的数据
    fun readFromMemberList(){
        val filePath = FILE_PATH
        val workbook: Workbook = try {
            // 读取已有文件
            FileInputStream(filePath).use { WorkbookFactory.create(it) }
        } catch (e: Exception) {
            // 文件不存在时创建新工作簿
            XSSFWorkbook()
        }

        // 获取已有sheet或创建新sheet
        val sheet = workbook.getSheet("人员信息") ?: workbook.createSheet("人员信息")

        //drop(1)代表跳过首行
        _memberList= (sheet?.drop(1)?.mapNotNull { row ->
            //尝试将单元格信息读取进memberList数组
            try {
                Member(
                    name = row.getCell(0).toString(), //人员姓名
                    id = row.getCell(1).toString() //钉钉ID
                )
            } catch (e: Exception) {
                null  // 忽略格式错误行
            }
        } ?: emptyList()).toMutableList()

        //初始化缓存映射
        _memberMap.putAll(memberList.associate { it.name to it.id })

        //关闭工作簿
        try {
            workbook.close()
        }catch (e: IOException){
            println(e)
        }
    }
}