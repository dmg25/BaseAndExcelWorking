using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Data;

using FirebirdSql.Data.FirebirdClient;


namespace ConnectToSCADABD
{
    // подключение к БД и выполнение одиночного запроса
    public class ProgramConnect
    {

        public DataTable dt1 = new System.Data.DataTable(); // таблица с данными из БД

        public void ConnectToBase(string SQL, string BaseAddr )
        {
            Form1 f = new Form1();  //тащим данные из формы
            string ConStr = "character set=WIN1251;initial catalog=" + BaseAddr + ";user id=SYSDBA;password=masterkey"; // наша строка подключения, сделать её изменяемой!!!

            // Строка подключения
            using (FbConnection fbc = new FbConnection(ConStr))   // используем Using для последующего высвобождения ресурсов, должно быть оптимальнее
            {
                try
                {
                    fbc.Open();
                }
                catch (Exception ex)
                {
                    MessageBox.Show(ex.Message,
                        System.IO.Path.GetFileName(
                        System.Reflection.Assembly.GetExecutingAssembly().Location));
                    return;
                }

                // Транзакция, которая тупо откуда-то скопирована
                FbTransactionOptions fbto = new FbTransactionOptions();
                fbto.TransactionBehavior = FbTransactionBehavior.NoWait |
                     FbTransactionBehavior.ReadCommitted |
                     FbTransactionBehavior.RecVersion;
                FbTransaction fbt = fbc.BeginTransaction(fbto);

                FbCommand fbcom = new FbCommand(SQL, fbc, fbt);

                // Создаем адаптер данных
                FbDataAdapter fbda = new FbDataAdapter(fbcom);

                DataSet ds = new DataSet();   //из формы не тащится, видимо там надо задавать способ подключения через стандартные способы, а они не пашут.

                try
                {
                    fbda.Fill(ds);  // заполняем DataSet
                }
                catch (Exception ex)
                {
                    MessageBox.Show(ex.Message,
                        System.IO.Path.GetFileName(
                        System.Reflection.Assembly.GetExecutingAssembly().Location));
                    return;
                }
                finally
                {
                    fbt.Rollback();
                    fbc.Close();
                }

                dt1 = ds.Tables[0];

                if (dt1.Rows.Count == 0) return;

            }
        }

        public void WriteToBase(string SQL, string BaseAddr)
        {
            Form1 f = new Form1();  //тащим данные из формы
            string ConStr = "character set=WIN1251;initial catalog=" + BaseAddr + ";user id=SYSDBA;password=masterkey"; // наша строка подключения, сделать её изменяемой!!!

            // Строка подключения
            using (FbConnection fbc = new FbConnection(ConStr))   // используем Using для последующего высвобождения ресурсов, должно быть оптимальнее
            {
                try
                {
                    fbc.Open();
                }
                catch (Exception ex)
                {
                    MessageBox.Show(ex.Message,
                        System.IO.Path.GetFileName(
                        System.Reflection.Assembly.GetExecutingAssembly().Location));
                    return;
                }

                // Транзакция, которая тупо откуда-то скопирована
              /*  FbTransactionOptions fbto = new FbTransactionOptions();
                fbto.TransactionBehavior = FbTransactionBehavior.NoWait |
                     FbTransactionBehavior.ReadCommitted |
                     FbTransactionBehavior.RecVersion;*/
                FbTransaction fbt = fbc.BeginTransaction(/*fbto*/);

               

                FbCommand fbcom = new FbCommand(SQL, fbc, fbt);

                fbcom.Transaction = fbt;

               // fbcom.ExecuteNonQuery();


                try
                {
                    int res = fbcom.ExecuteNonQuery(); //для запросов, не возвращающих набор данных (insert, update, delete) надо вызывать этот метод
                 //   MessageBox.Show("SUCCESS: " + res.ToString());
                    fbt.Commit(); //если вставка прошла успешно - комитим транзакцию
                }
                catch (Exception ex)
                {
                    MessageBox.Show(ex.Message);
                }

                    fbcom.Dispose(); //в документации написано, что ОЧЕНЬ рекомендуется убивать объекты этого типа, если они больше не нужны
                    //fbt.Rollback();
                    fbc.Close();
                

             





                // Создаем адаптер данных
             /*   FbDataAdapter fbda = new FbDataAdapter(fbcom);

                DataSet ds = new DataSet();   //из формы не тащится, видимо там надо задавать способ подключения через стандартные способы, а они не пашут.

                try
                {
                    fbda.Fill(ds);  // заполняем DataSet
                }
                catch (Exception ex)
                {
                    MessageBox.Show(ex.Message,
                        System.IO.Path.GetFileName(
                        System.Reflection.Assembly.GetExecutingAssembly().Location));
                    return;
                }
                finally
                {
                    fbt.Rollback();
                    fbc.Close();
                }

                dt1 = ds.Tables[0];

                if (dt1.Rows.Count == 0) return;*/

            }
        }
    }
}
