using ATTAS_API.Models;
using System;
using System.Data.SqlClient;

namespace ATTAS_API.Utils
{
    public class SqlServerConnector
    {
        private readonly string connectionString;

        public SqlServerConnector(string serverName, string dbName, string username, string password)
        {
            connectionString = $"Server={serverName};Database={dbName};User Id={username};Password={password};";
        }

        public int addSession(string sessionHash)
        {
            using (SqlConnection connection = new SqlConnection(connectionString))
            {
                try
                {
                    connection.Open();
                    string sql = "INSERT INTO [session] (sessionHash, statusId, solutionCount) OUTPUT INSERTED.ID VALUES (@val1, @val2, @val3)";
                    using (SqlCommand command = new SqlCommand(sql, connection))
                    {
                        command.Parameters.AddWithValue("@val1", sessionHash);
                        command.Parameters.AddWithValue("@val2", 1);
                        command.Parameters.AddWithValue("@val3", 0);

                        int insertedId = (int)command.ExecuteScalar();
                        connection.Close();
                        return insertedId;
                    }
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"Error SQL Server : {ex.Message}");
                    return -1;
                }
                finally
                {
                    connection.Close();
                }
            }
        }
        public int addTask(int sessionId, string businessId, int order)
        {
            using (SqlConnection connection = new SqlConnection(connectionString))
            {
                try
                {
                    connection.Open();
                    string sql = "INSERT INTO [task] (sessionId, businessId, [order]) OUTPUT INSERTED.ID VALUES (@val1, @val2, @val3)";
                    using (SqlCommand command = new SqlCommand(sql, connection))
                    {
                        command.Parameters.AddWithValue("@val1", sessionId);
                        command.Parameters.AddWithValue("@val2", businessId);
                        command.Parameters.AddWithValue("@val3", order);

                        int insertedId = (int)command.ExecuteScalar();
                        connection.Close();
                        return insertedId;
                    }
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"Error SQL Server : {ex.Message}");
                    return -1;
                }
                finally
                {
                    connection.Close();
                }
            }
        }

        public int addInstructor(int sessionId, string businessId, int order)
        {
            using (SqlConnection connection = new SqlConnection(connectionString))
            {
                try
                {
                    connection.Open();
                    string sql = "INSERT INTO [instructor] (sessionId, businessId, [order]) OUTPUT INSERTED.ID VALUES (@val1, @val2, @val3)";
                    using (SqlCommand command = new SqlCommand(sql, connection))
                    {
                        command.Parameters.AddWithValue("@val1", sessionId);
                        command.Parameters.AddWithValue("@val2", businessId);
                        command.Parameters.AddWithValue("@val3", order);

                        int insertedId = (int)command.ExecuteScalar();
                        connection.Close();
                        return insertedId;
                    }
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"Error SQL Server : {ex.Message}");
                    return -1;
                }
                finally
                {
                    connection.Close();
                }
            }
        }
        public int addTime(int sessionId, string businessId, int order)
        {
            using (SqlConnection connection = new SqlConnection(connectionString))
            {
                try
                {
                    connection.Open();
                    string sql = "INSERT INTO [time] (sessionId, businessId, [order]) OUTPUT INSERTED.ID VALUES (@val1, @val2, @val3)";
                    using (SqlCommand command = new SqlCommand(sql, connection))
                    {
                        command.Parameters.AddWithValue("@val1", sessionId);
                        command.Parameters.AddWithValue("@val2", businessId);
                        command.Parameters.AddWithValue("@val3", order);

                        int insertedId = (int)command.ExecuteScalar();
                        connection.Close();
                        return insertedId;
                    }
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"Error SQL Server : {ex.Message}");
                    return -1;
                }
                finally
                {
                    connection.Close();
                }
            }
        }
        public int updateSessionStatus(int sessionId, int statusId,int solutionCount)
        {
            using (SqlConnection connection = new SqlConnection(connectionString))
            {
                try
                {
                    connection.Open();
                    string updateQuery = "UPDATE session SET statusId=@value1, solutionCount=@value2 WHERE ID=@id";

                    using (SqlCommand command = new SqlCommand(updateQuery, connection))
                    {
                        command.Parameters.AddWithValue("@value1", statusId);
                        command.Parameters.AddWithValue("@value2", solutionCount);
                        command.Parameters.AddWithValue("@id", sessionId);

                        // Execute the query and get the number of affected rows
                        int rowsAffected = command.ExecuteNonQuery();
                        connection.Close();
                        return rowsAffected;
                    }
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"Error SQL Server : {ex.Message}");
                    return 0;
                }
                finally
                {
                    connection.Close();
                }
            }
        }
        public int addSolution(int sessionId, int no,int taskAssigned,int workingDay,int workingTime,int waitingTime,int subjectDiversity,int quotaAvailabe,int walkingDistance,int subjectPreference,int slotPreference)
        {
            using (SqlConnection connection = new SqlConnection(connectionString))
            {
                try
                {
                    connection.Open();
                    string sql = "INSERT INTO [solution] (sessionId,no,taskAssigned,workingDay,workingTime,waitingTime,subjectDiversity,quotaAvailable,walkingDistance,subjectPreference,slotPreference) OUTPUT INSERTED.ID VALUES (@val0, @val1, @val2, @val3, @val4, @val5, @val6, @val7, @val8, @val9, @val10 )";
                    using (SqlCommand command = new SqlCommand(sql, connection))
                    {
                        command.Parameters.AddWithValue("@val0", sessionId);
                        command.Parameters.AddWithValue("@val1", no);
                        command.Parameters.AddWithValue("@val2", taskAssigned);
                        command.Parameters.AddWithValue("@val3", workingDay);
                        command.Parameters.AddWithValue("@val4", workingTime);
                        command.Parameters.AddWithValue("@val5", waitingTime);
                        command.Parameters.AddWithValue("@val6", subjectDiversity);
                        command.Parameters.AddWithValue("@val7", quotaAvailabe);
                        command.Parameters.AddWithValue("@val8", walkingDistance);
                        command.Parameters.AddWithValue("@val9", subjectPreference);
                        command.Parameters.AddWithValue("@val10", slotPreference);

                        int insertedId = (int)command.ExecuteScalar();
                        connection.Close();
                        return insertedId;
                    }
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"Error SQL Server : {ex.Message}");
                    return -1;
                }
                finally
                {
                    connection.Close();
                }
            }
        }
        public int addResult(int solutionId,int taskId,int instructorId,int timeId)
        {
            using (SqlConnection connection = new SqlConnection(connectionString))
            {
                try
                {
                    connection.Open();
                    string sql = "INSERT INTO [result] (solutionId,taskOrder,instructorOrder,timeOrder) OUTPUT INSERTED.ID VALUES (@val0, @val1, @val2, @val3)";
                    using (SqlCommand command = new SqlCommand(sql, connection))
                    {
                        command.Parameters.AddWithValue("@val0", solutionId);
                        command.Parameters.AddWithValue("@val1", taskId);
                        command.Parameters.AddWithValue("@val2", instructorId);
                        command.Parameters.AddWithValue("@val3", timeId);

                        int insertedId = (int)command.ExecuteScalar();
                        connection.Close();
                        return insertedId;
                    }
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"Error SQL Server : {ex.Message}");
                    return -1;
                }
                finally
                {
                    connection.Close();
                }
            }
        }
        public bool validToken(string token)
        {
            using (SqlConnection connection = new SqlConnection(connectionString))
            {
                try
                {
                    connection.Open();
                    string sql = "SELECT * FROM [token] WHERE tokenHash=@token";
                    using (SqlCommand command = new SqlCommand(sql, connection))
                    {
                        command.Parameters.AddWithValue("@token", token);
                        SqlDataReader reader = command.ExecuteReader();
                        while (reader.Read())
                        {
                            connection.Close();
                            return true;
                        }
                        reader.Close();
                    }
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"Error SQL Server : {ex.Message}");
                    return false;
                }
                finally
                {
                    connection.Close();
                }
                return false;
            }
        }
    }
}
