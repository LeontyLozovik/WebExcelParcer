using System.ComponentModel.DataAnnotations.Schema;
using System.ComponentModel.DataAnnotations;
using System.Diagnostics.Metrics;
using System.ComponentModel;

namespace WebParcer.Models.TableModels
{
    public class TableModelBase     //базовый макет для таблиц базы данных
    {
        public int Id { get; set; }

        [DisplayName("Б/сч")]
        public int B_sch { get; set; }

        [DisplayName("Входящее сальдо актив")]
        public double InBalanceActive { get; set; }

        [DisplayName("Входящее сальдо пасив")]
        public double InBalancePassive { get; set; }

        [DisplayName("Дебит")]
        public double Debit { get; set; }

        [DisplayName("Кредит")]
        public double Credit { get; set; }

        [DisplayName("Исходящее сальдо актив")]
        public double OutBalanceActive { get; set; }

        [DisplayName("Исходящее сальдо пасив")]
        public double OutBalancePassive { get; set; }
    }
}
