//! Validate inline user attachments before sending them to ACP.
use agent_client_protocol::schema::v1::{ContentBlock, ImageContent, TextContent};
use base64::{engine::general_purpose::STANDARD, Engine};
use serde::Deserialize;

#[derive(Deserialize)]
#[serde(rename_all = "camelCase")]
pub struct PromptImage {
    pub data: String,
    pub mime_type: String,
}

pub fn blocks(text: &str, images: Vec<PromptImage>) -> Result<Vec<ContentBlock>, String> {
    if images.len() > 4 {
        return Err("每轮最多发送 4 张图片".into());
    }
    let mut content = vec![ContentBlock::Text(TextContent::new(text))];
    for image in images {
        if !matches!(
            image.mime_type.as_str(),
            "image/png" | "image/jpeg" | "image/webp"
        ) {
            return Err("图片格式仅支持 PNG、JPEG、WebP".into());
        }
        if image.data.len() > 7_000_000 {
            return Err("每张图片不得超过 5MB".into());
        }
        let bytes = STANDARD
            .decode(&image.data)
            .map_err(|_| "图片数据无效，请重新添加附件")?;
        if bytes.is_empty() || bytes.len() > 5 * 1024 * 1024 {
            return Err("图片必须非空且不超过 5MB".into());
        }
        content.push(ContentBlock::Image(ImageContent::new(
            image.data,
            image.mime_type,
        )));
    }
    Ok(content)
}

#[cfg(test)]
mod tests {
    use super::*;
    #[test]
    fn image_blocks_preserve_mime_and_reject_invalid_payloads() {
        let image = || PromptImage {
            data: "aGVsbG8=".into(),
            mime_type: "image/png".into(),
        };
        let value = serde_json::to_value(blocks("图像", vec![image()]).unwrap()).unwrap();
        assert_eq!(value[0]["text"], "图像");
        assert_eq!(value[1]["type"], "image");
        assert_eq!(value[1]["mimeType"], "image/png");
        assert!(blocks(
            "",
            vec![PromptImage {
                data: "%%%".into(),
                ..image()
            }]
        )
        .is_err());
        assert!(blocks(
            "",
            vec![PromptImage {
                mime_type: "text/html".into(),
                ..image()
            }]
        )
        .is_err());
        assert!(blocks("", (0..5).map(|_| image()).collect()).is_err());
    }
}
