﻿namespace Fussballteam.Models
{
    public class ErrorView
    {
        public string? RequestId { get; set; }

        public bool ShowRequestId => !string.IsNullOrEmpty(RequestId);
    }
}