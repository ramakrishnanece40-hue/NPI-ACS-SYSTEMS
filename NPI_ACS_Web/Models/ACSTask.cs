using System;
using System.ComponentModel.DataAnnotations;

namespace NPI_ACS_Web.Models
{
    public class ACSTask
    {
        public int Id { get; set; }

        public string? Project { get; set; }
        public string? ODM { get; set; }
        public string? Product { get; set; }
        public string? Model { get; set; }
        public string? Question { get; set; }

        public string? ActionDetail { get; set; }
        public string? FourM { get; set; }

        public string? NeolyncPIC { get; set; }
        public string? CustomerPIC { get; set; }

        public string? Priority { get; set; }
        public string? Status { get; set; }

        [DataType(DataType.Date)]
        public DateOnly? StartDate { get; set; }

        [DataType(DataType.Date)]
        public DateOnly? DueDate { get; set; }

        [DataType(DataType.Date)]
        public DateOnly? ActualCloseDate { get; set; }

        public string? Remarks { get; set; }

        public string? AttachmentPath { get; set; }
    }
}