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
// ������ �������� ��
System::Void Laba60::MyForm::button_download_Click(System::Object^ sender, System::EventArgs^ e)
{
	//������ ����������� ��
	String^ connectionString = "Provider=Microsoft.ACE.OLEDB.12.0; Data Source=123.accdb"; // ������ �����������
	OleDbConnection^ dbConnection = gcnew OleDbConnection(connectionString);

	//��������� ������ � ��
	dbConnection->Open(); // ��������� ����������
	String^ query = "SELECT * FROM [table_name]"; // ������
	OleDbCommand^ dbComand = gcnew OleDbCommand(query, dbConnection); // �������
	OleDbDataReader^ dbReader = dbComand->ExecuteReader(); // ��������� ������

	// ��������� ��������� ������
	if (dbReader->HasRows == false)
	{
		MessageBox::Show("������ ���������� ������", "������!");
	}
	else
	{
		while (dbReader->Read())
		{
			dataGridView1->Rows->Add(dbReader["id"], dbReader["Name"], dbReader["Price"], dbReader["Quantity"]);
		}
	}

	// ��������� ����������
	dbReader->Close();
	dbConnection->Close();

	return System::Void();
}

// ������ ���������� � ��
System::Void Laba60::MyForm::button_add_Click(System::Object^ sender, System::EventArgs^ e)
{
	// ����� ������ ������ ��� ����������
	if (dataGridView1->SelectedRows->Count != 1)
	{
		MessageBox::Show("�������� ���� ������ ��� ����������!", "��������!");
		return;
	}

	// ������ ������ ��������� ������
	int index = dataGridView1->SelectedRows[0]->Index;

	// ��������� ������
	if (dataGridView1->Rows[index]->Cells[0]->Value == nullptr ||
		dataGridView1->Rows[index]->Cells[1]->Value == nullptr || 
		dataGridView1->Rows[index]->Cells[2]->Value == nullptr || 
		dataGridView1->Rows[index]->Cells[3]->Value == nullptr)
	{
		MessageBox::Show("�� ��� ������ ��������!", "���������!");
		return;
	}

	// ��������� ������
	String^ id = dataGridView1->Rows[index]->Cells[0]->Value->ToString();
	String^ name = dataGridView1->Rows[index]->Cells[1]->Value->ToString();
	String^ Price = dataGridView1->Rows[index]->Cells[2]->Value->ToString();
	String^ Quantity = dataGridView1->Rows[index]->Cells[3]->Value->ToString();


	//������ ����������� ��
	String^ connectionString = "Provider=Microsoft.ACE.OLEDB.12.0; Data Source=123.accdb"; // ������ �����������
	OleDbConnection^ dbConnection = gcnew OleDbConnection(connectionString);

	//��������� ������ � ��
	dbConnection->Open(); // ��������� ����������
	String^ query = "INSERT INTO [table_name] VALUES (" + id + ", '" + name + "', " + Price + ", " + Quantity + ")"; // ������
	OleDbCommand^ dbComand = gcnew OleDbCommand(query, dbConnection); // �������
	
	// ��������� ������
	if (dbComand->ExecuteNonQuery() != 1)
	{
		MessageBox::Show("������ ���������� �������!", "������!");
	}
	else
		MessageBox::Show("������ ���������", " ������!");

	// ��������� ����������
	dbConnection->Close();

	return System::Void();
}

// ������ ���������� ������ � ������� � ��
System::Void Laba60::MyForm::button_update_Click(System::Object^ sender, System::EventArgs^ e)
{

	// ����� ������ ������ ��� ����������
	if (dataGridView1->SelectedRows->Count != 1)
	{
		MessageBox::Show("�������� ���� ������ ��� ����������!", "��������!");
		return;
	}

	// ������ ������ ��������� ������
	int index = dataGridView1->SelectedRows[0]->Index;

	// ��������� ������
	if (dataGridView1->Rows[index]->Cells[0]->Value == nullptr ||
		dataGridView1->Rows[index]->Cells[1]->Value == nullptr ||
		dataGridView1->Rows[index]->Cells[2]->Value == nullptr ||
		dataGridView1->Rows[index]->Cells[3]->Value == nullptr)
	{
		MessageBox::Show("�� ��� ������ ��������!", "���������!");
		return;
	}

	// ��������� ������
	String^ id = dataGridView1->Rows[index]->Cells[0]->Value->ToString();
	String^ name = dataGridView1->Rows[index]->Cells[1]->Value->ToString();
	String^ Price = dataGridView1->Rows[index]->Cells[2]->Value->ToString();
	String^ Quantity = dataGridView1->Rows[index]->Cells[3]->Value->ToString();

	//������ ����������� ��
	String^ connectionString = "Provider=Microsoft.ACE.OLEDB.12.0; Data Source=123.accdb"; // ������ �����������
	OleDbConnection^ dbConnection = gcnew OleDbConnection(connectionString);

	//��������� ������ � ��
	dbConnection->Open(); // ��������� ����������
	String^ query = "UPDATE [table_name] SET Name = '"+ name +"', Price = "+ Price +", Quantity = "+ Quantity +" WHERE id = " + id;// ������
	OleDbCommand^ dbComand = gcnew OleDbCommand(query, dbConnection); // �������

	// ��������� ������
	if (dbComand->ExecuteNonQuery() != 1)
	{
		MessageBox::Show("������ ���������� �������!", "������!");
	}
	else
		MessageBox::Show("������ ���������!", " ������!");

	// ��������� ����������
	dbConnection->Close();

	return System::Void();
}

// ������ ��� �������� ������ �� ������� � ��
System::Void Laba60::MyForm::button_delete_Click(System::Object^ sender, System::EventArgs^ e)
{
	// ����� ������ ������ ��� ����������
	if (dataGridView1->SelectedRows->Count != 1)
	{
		MessageBox::Show("�������� ���� ������ ��� ����������!", "��������!");
		return;
	}

	// ������ ������ ��������� ������
	int index = dataGridView1->SelectedRows[0]->Index;

	// ��������� ������
	if (dataGridView1->Rows[index]->Cells[0]->Value == nullptr)
	{
		MessageBox::Show("�� ��� ������ ��������!", "��������!");
		return;
	}

	// ��������� ������
	String^ id = dataGridView1->Rows[index]->Cells[0]->Value->ToString();

	//������ ����������� ��
	String^ connectionString = "Provider=Microsoft.ACE.OLEDB.12.0; Data Source=123.accdb"; // ������ �����������
	OleDbConnection^ dbConnection = gcnew OleDbConnection(connectionString);

	//��������� ������ � ��
	dbConnection->Open(); // ��������� ����������
	String^ query = " DELETE FROM [table_name] WHERE id = " + id;// ������
	OleDbCommand^ dbComand = gcnew OleDbCommand(query, dbConnection); // �������

	// ��������� ������
	if (dbComand->ExecuteNonQuery() != 1)
	{
		MessageBox::Show("������ ���������� �������!", "������!");	
	}
	else {
		MessageBox::Show("������ �������!", " ������!");
		dataGridView1->Rows->RemoveAt(index); // ������� ������ �� ������� ��
	}
	// ��������� ����������
	dbConnection->Close();


	return System::Void();
}
