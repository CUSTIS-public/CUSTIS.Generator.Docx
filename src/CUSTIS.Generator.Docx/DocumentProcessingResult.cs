﻿using Newtonsoft.Json.Linq;

namespace CUSTIS.DocumentService;
public class DocumentProcessingResult
{
    public bool Success { get; init; }
    public JObject Result { get; init; }
    public DocumentProcessingResult(bool success, JObject result) { Success = success; Result = result; }
}