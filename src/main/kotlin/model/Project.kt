package model

import java.time.LocalDate

data class Project(
    val name:String,
    val time: LocalDate,
    val message:String,
    var isMessageSend:Boolean=false,
    val memberName:String?
)