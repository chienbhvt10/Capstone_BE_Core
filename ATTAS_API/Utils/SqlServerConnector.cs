using ATTAS_API.Models;
using OperationsResearch;
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
        public Session getSession(string hash)
        {
            Session session = new Session();
            using (SqlConnection connection = new SqlConnection(connectionString))
            {
                try
                {
                    connection.Open();
                    string sql = "SELECT * FROM [session] WHERE sessionHash=@hash";
                    using (SqlCommand command = new SqlCommand(sql, connection))
                    {
                        command.Parameters.AddWithValue("@hash", hash);
                        SqlDataReader reader = command.ExecuteReader();
                        while (reader.Read())
                        {
                            session.id = (int)reader[0];
                            session.hash = (string)reader[1];
                            session.statusId = (int)reader[2];
                            session.solutionCount = (int)reader[3];
                            connection.Close();
                            return session;
                        }
                        reader.Close();
                    }
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"Error SQL Server : {ex.Message}");
                    return null;
                }
                finally
                {
                    connection.Close();
                }
                return null;
            }
        }
        public Solution getSolution(int sessionId,int no)
        {
            Solution solution = new Solution();
            using (SqlConnection connection = new SqlConnection(connectionString))
            {
                try
                {
                    connection.Open();
                    string sql = "SELECT * FROM [solution] WHERE sessionId=@sessionid AND no=@no";
                    using (SqlCommand command = new SqlCommand(sql, connection))
                    {
                        command.Parameters.AddWithValue("@sessionid", sessionId);
                        command.Parameters.AddWithValue("@no", no);
                        SqlDataReader reader = command.ExecuteReader();
                        while (reader.Read())
                        {
                            solution.Id = (int)reader[0];
                            solution.sessionId = (int)reader[1];
                            solution.no = (int)reader[2];
                            solution.taskAssigned = (int)reader[3];
                            solution.workingDay = (int)reader[4];
                            solution.workingTime = (int)reader[5];
                            solution.waitingTime = (int)reader[6];
                            solution.subjectDiversity = (int)reader[7];
                            solution.quotaAvailable= (int)reader[8];
                            solution.walkingDistance = (int)reader[9];
                            solution.subjectPreference = (int)reader[10];
                            solution.slotPreference = (int)reader[11];
                            connection.Close();
                            return solution;
                        }
                        reader.Close();
                    }
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"Error SQL Server : {ex.Message}");
                    return null;
                }
                finally
                {
                    connection.Close();
                }
                return null;
            }
        }
        public List<Assigned> getResult(int solutionId,int sessionId)
        {
            List<Assigned> assigneds = new List<Assigned>();
            using (SqlConnection connection = new SqlConnection(connectionString))
            {
                try
                {
                    connection.Open();
                    string sql = "SELECT * FROM [result] WHERE solutionId=@solutionid";
                    using (SqlCommand command = new SqlCommand(sql, connection))
                    {
                        command.Parameters.AddWithValue("@solutionid", solutionId);
                        SqlDataReader reader = command.ExecuteReader();
                        while (reader.Read())
                        {
                            Assigned assigned = new Assigned();
                            assigned.taskId = getBusinessId("task", sessionId, (int)reader[2]);
                            assigned.instructorId = getBusinessId("instructor",sessionId, (int)reader[3]);
                            assigned.slotId = getBusinessId("time", sessionId, (int)reader[4]);
                            assigneds.Add(assigned);
                            
                        }
                        reader.Close();
                        connection.Close();
                        return assigneds;
                    }
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"Error SQL Server : {ex.Message}");
                    return null;
                }
                finally
                {
                    connection.Close();
                }
                return null;
            }
        }
        public string getBusinessId(string table,int sessionId,int order)
        {
            using (SqlConnection connection = new SqlConnection(connectionString))
            {
                try
                {
                    connection.Open();
                    string sql = $"SELECT * FROM [{table}] WHERE sessionId=@sessionid AND [order]=@order";
                    using (SqlCommand command = new SqlCommand(sql, connection))
                    {
                        command.Parameters.AddWithValue("@sessionid", sessionId);
                        command.Parameters.AddWithValue("@order", order);
                        SqlDataReader reader = command.ExecuteReader();
                        while (reader.Read())
                        {
                            string tmp = (string)reader[2];
                            reader.Close();
                            connection.Close();
                            return tmp;
                        }
                        
                    }
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"Error SQL Server : {ex.Message}");
                    return null;
                }
                finally
                {
                    connection.Close();
                }
                return null;
            }
        }
    }
}
