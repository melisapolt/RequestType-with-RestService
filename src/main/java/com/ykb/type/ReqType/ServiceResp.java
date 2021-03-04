package com.ykb.type.ReqType;

import lombok.Data;

@Data
public class ServiceResp {
private InfraStructureError error;
private String sendStatus;
private Status status;
}
