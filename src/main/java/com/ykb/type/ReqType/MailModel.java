package com.ykb.type.ReqType;

import lombok.Data;

@Data
public class MailModel {
	private String analyst;
	private String application;
	private Attachment[] attachments;
	private String ccList;
	private String content;
	private String fromAddress;
	private String fromName;
	private String replyToAddress;
	private String replyToName;
	private String subject;
	private String toList;
	private String user;
}
