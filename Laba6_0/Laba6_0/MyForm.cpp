#include "MyForm.h"

using namespace System;
using namespace System::Windows::Forms;
using namespace System::Data::OleDb;

[STAThreadAttribute]
int main(array<String^>^ args) {
	Application::SetCompatibleTextRenderingDefault(false);
	Application::EnableVisualStyles();

	Laba60::MyForm form;
	Application::Run(% form);
}
// кнопка загрузки БД
System::Void Laba60::MyForm::button_download_Click(System::Object^ sender, System::EventArgs^ e)
{
	//Строка подключения БД
	String^ connectionString = "Provider=Microsoft.ACE.OLEDB.12.0; Data Source=123.accdb"; // строка подключения
	OleDbConnection^ dbConnection = gcnew OleDbConnection(connectionString);

	//Выполняем заброс к БД
	dbConnection->Open(); // открываем соединение
	String^ query = "SELECT * FROM [table_name]"; // запрос
	OleDbCommand^ dbComand = gcnew OleDbCommand(query, dbConnection); // команда
	OleDbDataReader^ dbReader = dbComand->ExecuteReader(); // считываем данные

	// проверяем введенные данные
	if (dbReader->HasRows == false)
	{
		MessageBox::Show("Ошибка считывания данных", "Ошибка!");
	}
	else
	{
		while (dbReader->Read())
		{
			dataGridView1->Rows->Add(dbReader["id"], dbReader["Name"], dbReader["Price"], dbReader["Quantity"]);
		}
	}

	// закрываем соединение
	dbReader->Close();
	dbConnection->Close();

	return System::Void();
}

// кнопка довавления в БД
System::Void Laba60::MyForm::button_add_Click(System::Object^ sender, System::EventArgs^ e)
{
	// выбор нужной строки для добавления
	if (dataGridView1->SelectedRows->Count != 1)
	{
		MessageBox::Show("Выберите одну строку для добавления!", "Внимание!");
		return;
	}

	// узнаем индекс выбранной строки
	int index = dataGridView1->SelectedRows[0]->Index;

	// проверяем данные
	if (dataGridView1->Rows[index]->Cells[0]->Value == nullptr ||
		dataGridView1->Rows[index]->Cells[1]->Value == nullptr || 
		dataGridView1->Rows[index]->Cells[2]->Value == nullptr || 
		dataGridView1->Rows[index]->Cells[3]->Value == nullptr)
	{
		MessageBox::Show("Не все данные введенны!", "Вниманиек!");
		return;
	}

	// считываем данные
	String^ id = dataGridView1->Rows[index]->Cells[0]->Value->ToString();
	String^ name = dataGridView1->Rows[index]->Cells[1]->Value->ToString();
	String^ Price = dataGridView1->Rows[index]->Cells[2]->Value->ToString();
	String^ Quantity = dataGridView1->Rows[index]->Cells[3]->Value->ToString();


	//Строка подключения БД
	String^ connectionString = "Provider=Microsoft.ACE.OLEDB.12.0; Data Source=123.accdb"; // строка подключения
	OleDbConnection^ dbConnection = gcnew OleDbConnection(connectionString);

	//Выполняем заброс к БД
	dbConnection->Open(); // открываем соединение
	String^ query = "INSERT INTO [table_name] VALUES (" + id + ", '" + name + "', " + Price + ", " + Quantity + ")"; // запрос
	OleDbCommand^ dbComand = gcnew OleDbCommand(query, dbConnection); // команда
	
	// Выполняем запрос
	if (dbComand->ExecuteNonQuery() != 1)
	{
		MessageBox::Show("Ошибка выполнения запроса!", "Ошибка!");
	}
	else
		MessageBox::Show("Данные добавлены", " Готово!");

	// закрываем соединение
	dbConnection->Close();

	return System::Void();
}

// кнопка обновления данных в таблице и БД
System::Void Laba60::MyForm::button_update_Click(System::Object^ sender, System::EventArgs^ e)
{

	// выбор нужной строки для добавления
	if (dataGridView1->SelectedRows->Count != 1)
	{
		MessageBox::Show("Выберите одну строку для добавления!", "Внимание!");
		return;
	}

	// узнаем индекс выбранной строки
	int index = dataGridView1->SelectedRows[0]->Index;

	// проверяем данные
	if (dataGridView1->Rows[index]->Cells[0]->Value == nullptr ||
		dataGridView1->Rows[index]->Cells[1]->Value == nullptr ||
		dataGridView1->Rows[index]->Cells[2]->Value == nullptr ||
		dataGridView1->Rows[index]->Cells[3]->Value == nullptr)
	{
		MessageBox::Show("Не все данные введенны!", "Вниманиек!");
		return;
	}

	// считываем данные
	String^ id = dataGridView1->Rows[index]->Cells[0]->Value->ToString();
	String^ name = dataGridView1->Rows[index]->Cells[1]->Value->ToString();
	String^ Price = dataGridView1->Rows[index]->Cells[2]->Value->ToString();
	String^ Quantity = dataGridView1->Rows[index]->Cells[3]->Value->ToString();

	//Строка подключения БД
	String^ connectionString = "Provider=Microsoft.ACE.OLEDB.12.0; Data Source=123.accdb"; // строка подключения
	OleDbConnection^ dbConnection = gcnew OleDbConnection(connectionString);

	//Выполняем заброс к БД
	dbConnection->Open(); // открываем соединение
	String^ query = "UPDATE [table_name] SET Name = '"+ name +"', Price = "+ Price +", Quantity = "+ Quantity +" WHERE id = " + id;// запрос
	OleDbCommand^ dbComand = gcnew OleDbCommand(query, dbConnection); // команда

	// Выполняем запрос
	if (dbComand->ExecuteNonQuery() != 1)
	{
		MessageBox::Show("Ошибка выполнения запроса!", "Ошибка!");
	}
	else
		MessageBox::Show("Данные измененны!", " Готово!");

	// закрываем соединение
	dbConnection->Close();

	return System::Void();
}

// кнопка для удаления данных из таблицы и БД
System::Void Laba60::MyForm::button_delete_Click(System::Object^ sender, System::EventArgs^ e)
{
	// выбор нужной строки для добавления
	if (dataGridView1->SelectedRows->Count != 1)
	{
		MessageBox::Show("Выберите одну строку для добавления!", "Внимание!");
		return;
	}

	// узнаем индекс выбранной строки
	int index = dataGridView1->SelectedRows[0]->Index;

	// проверяем данные
	if (dataGridView1->Rows[index]->Cells[0]->Value == nullptr)
	{
		MessageBox::Show("Не все данные введенны!", "Внимание!");
		return;
	}

	// считываем данные
	String^ id = dataGridView1->Rows[index]->Cells[0]->Value->ToString();

	//Строка подключения БД
	String^ connectionString = "Provider=Microsoft.ACE.OLEDB.12.0; Data Source=123.accdb"; // строка подключения
	OleDbConnection^ dbConnection = gcnew OleDbConnection(connectionString);

	//Выполняем заброс к БД
	dbConnection->Open(); // открываем соединение
	String^ query = " DELETE FROM [table_name] WHERE id = " + id;// запрос
	OleDbCommand^ dbComand = gcnew OleDbCommand(query, dbConnection); // команда

	// Выполняем запрос
	if (dbComand->ExecuteNonQuery() != 1)
	{
		MessageBox::Show("Ошибка выполнения запроса!", "Ошибка!");	
	}
	else {
		MessageBox::Show("Данные удалены!", " Готово!");
		dataGridView1->Rows->RemoveAt(index); // удаляем строку из таблицы БД
	}
	// закрываем соединение
	dbConnection->Close();


	return System::Void();
}
