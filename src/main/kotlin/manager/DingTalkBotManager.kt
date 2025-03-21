package manager

import com.dingtalk.api.DefaultDingTalkClient
import com.dingtalk.api.DingTalkClient
import com.dingtalk.api.request.OapiRobotSendRequest
import com.taobao.api.ApiException
import java.util.Base64
import javax.crypto.Mac
import javax.crypto.spec.SecretKeySpec
import java.io.UnsupportedEncodingException
import java.net.URLEncoder
import java.security.InvalidKeyException
import java.security.NoSuchAlgorithmException

object DingTalkBotManager {
    private const val CUSTOM_ROBOT_TOKEN = "94226b87fa780d6b37de71ef4ef8e40738710fae0b7a8ae538ba36e0649f538d"
    private const val SECRET = "SECc0faa607e7b1edc4971f1c17fc6092d95df3d18b5dd7fa5bf6a8e53e591090c2"

    var userIds: List<String> = emptyList() //需要@的用户ID
    var name:String="" //需发送的标题
    var message:String="" //需要发送的信息


    //发送Markdown格式的消息
    //@JvmStatic
    fun sendMarkdown() {
        try {
            val timestamp = System.currentTimeMillis()
            //println(timestamp)
            val stringToSign = "$timestamp\n$SECRET"
            val mac = Mac.getInstance("HmacSHA256").apply {
                init(SecretKeySpec(SECRET.toByteArray(Charsets.UTF_8), "HmacSHA256"))
            }
            val signData = mac.doFinal(stringToSign.toByteArray(Charsets.UTF_8))
            val sign = URLEncoder.encode(
                Base64.getEncoder().encodeToString(signData),
                Charsets.UTF_8.name()
            )
            //println("Debug Sign: $sign")
            val client: DingTalkClient = DefaultDingTalkClient(
                "https://oapi.dingtalk.com/robot/send?sign=$sign&timestamp=$timestamp"
            )
            val req = OapiRobotSendRequest().apply {
                //Markdown模式下无法@
                msgtype = "markdown"
                // 明确调用方法消除歧义
                setMarkdown(
                    OapiRobotSendRequest.Markdown().apply {
                        title= name
                        text= message
                    }
                )
            }
            val rsp = client.execute(req, CUSTOM_ROBOT_TOKEN)
            println(rsp.body)

        } catch (e: ApiException) {
            e.printStackTrace()
        } catch (e: Exception) {
            when (e) {
                is UnsupportedEncodingException,
                is NoSuchAlgorithmException,
                is InvalidKeyException -> throw RuntimeException(e)
                else -> throw e
            }
        }
    }


    //发送@消息
    fun sendAt(){
        try {
            val timestamp = System.currentTimeMillis()
            println(timestamp)
            val stringToSign = "$timestamp\n$SECRET"
            val mac = Mac.getInstance("HmacSHA256").apply {
                init(SecretKeySpec(SECRET.toByteArray(Charsets.UTF_8), "HmacSHA256"))
            }
            val signData = mac.doFinal(stringToSign.toByteArray(Charsets.UTF_8))
            val sign = URLEncoder.encode(
                Base64.getEncoder().encodeToString(signData),
                Charsets.UTF_8.name()
            )
            println("Debug Sign: $sign")
            val client: DingTalkClient = DefaultDingTalkClient(
                "https://oapi.dingtalk.com/robot/send?sign=$sign&timestamp=$timestamp"
            )
            val req = OapiRobotSendRequest().apply {
                msgtype = "text"
                // 明确调用方法消除歧义
                setText(
                    OapiRobotSendRequest.Text().apply {
                        content = "相关人员"
                    }
                )
                setAt(
                    OapiRobotSendRequest.At().apply {
                        atUserIds = userIds
                    }
                )
            }
            val rsp = client.execute(req, CUSTOM_ROBOT_TOKEN)
            println(rsp.body)

        } catch (e: ApiException) {
            e.printStackTrace()
        } catch (e: Exception) {
            when (e) {
                is UnsupportedEncodingException,
                is NoSuchAlgorithmException,
                is InvalidKeyException -> throw RuntimeException(e)
                else -> throw e
            }
        }
    }
}
