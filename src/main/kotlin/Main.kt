import manager.DingTalkBotManager
import manager.ExcelManager
import java.time.LocalDate

fun main(){
    println("程序开始执行")

    ExcelManager.init()
    ExcelManager.readFromProjectList()
    ExcelManager.readFromFinishedList()
    ExcelManager.readFromMemberList()

    for (project in ExcelManager.projectList) {
        if (!project.isMessageSend && project.time == LocalDate.now()) {
            println("\n发送项目: ${project.name}")

            // [1] 发送 Markdown 消息
            DingTalkBotManager.name = project.name
            DingTalkBotManager.message = "## ${project.name}\n##### ${project.message}"
            DingTalkBotManager.sendMarkdown()
            println("消息发送完成")

            Thread.sleep(500)

            // [2] 处理多用户@ (关键改进点)
            val targetIds = project.memberName
                ?.split(Regex("[、,，]")) // 通过正则表达式支持多种分隔符
                ?.mapNotNull { name -> ExcelManager.memberMap[name] } //利用缓存映射进行对比，减少时间复杂度
                ?.distinct() //去除集合里的重复元素，和Set不同的是它会保留顺序
                ?: emptyList()
            if (targetIds.isNotEmpty()) {
                DingTalkBotManager.userIds = targetIds
                DingTalkBotManager.sendAt()
                println("已通知相关人员")
                Thread.sleep(5600) //钉钉每分钟限制20条消息，包含@的消息总共有两条，故这里暂停约六秒
            }else{
                Thread.sleep(2600) //钉钉每分钟限制20条消息，不包含@的消息有一条，故这里暂停约三秒
            }

            // [3] 更新标记位
            project.isMessageSend = true
        }
    }

    ExcelManager.updateData()
    ExcelManager.storeToProjectSheet()
    ExcelManager.storeToFinishedSheet()

    println("\n程序执行完毕")
}