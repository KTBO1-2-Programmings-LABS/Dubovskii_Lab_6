#pragma once

namespace Laba60 {

	using namespace System;
	using namespace System::ComponentModel;
	using namespace System::Collections;
	using namespace System::Windows::Forms;
	using namespace System::Data;
	using namespace System::Drawing;

	/// <summary>
	/// Сводка для MyForm
	/// </summary>
	public ref class MyForm : public System::Windows::Forms::Form
	{
	public:
		MyForm(void)
		{
			InitializeComponent();
			//
			//TODO: добавьте код конструктора
			//
		}

	protected:
		/// <summary>
		/// Освободить все используемые ресурсы.
		/// </summary>
		~MyForm()
		{
			if (components)
			{
				delete components;
			}
		}
	private: System::Windows::Forms::DataGridView^ dataGridView1;
	private: System::Windows::Forms::GroupBox^ groupBox1;
	private: System::Windows::Forms::Button^ button_download;




	private: System::Windows::Forms::DataGridViewTextBoxColumn^ id;
	private: System::Windows::Forms::DataGridViewTextBoxColumn^ name;
	private: System::Windows::Forms::DataGridViewTextBoxColumn^ price;
	private: System::Windows::Forms::Button^ button_delete;

	private: System::Windows::Forms::Button^ button_update;

	private: System::Windows::Forms::Button^ button_add;
	private: System::Windows::Forms::DataGridViewTextBoxColumn^ quantity;

	protected:

	protected:

	private:
		/// <summary>
		/// Обязательная переменная конструктора.
		/// </summary>
		System::ComponentModel::Container ^components;

#pragma region Windows Form Designer generated code
		/// <summary>
		/// Требуемый метод для поддержки конструктора — не изменяйте 
		/// содержимое этого метода с помощью редактора кода.
		/// </summary>
		void InitializeComponent(void)
		{
			this->dataGridView1 = (gcnew System::Windows::Forms::DataGridView());
			this->id = (gcnew System::Windows::Forms::DataGridViewTextBoxColumn());
			this->name = (gcnew System::Windows::Forms::DataGridViewTextBoxColumn());
			this->price = (gcnew System::Windows::Forms::DataGridViewTextBoxColumn());
			this->quantity = (gcnew System::Windows::Forms::DataGridViewTextBoxColumn());
			this->groupBox1 = (gcnew System::Windows::Forms::GroupBox());
			this->button_delete = (gcnew System::Windows::Forms::Button());
			this->button_update = (gcnew System::Windows::Forms::Button());
			this->button_add = (gcnew System::Windows::Forms::Button());
			this->button_download = (gcnew System::Windows::Forms::Button());
			(cli::safe_cast<System::ComponentModel::ISupportInitialize^>(this->dataGridView1))->BeginInit();
			this->groupBox1->SuspendLayout();
			this->SuspendLayout();
			// 
			// dataGridView1
			// 
			this->dataGridView1->ColumnHeadersHeightSizeMode = System::Windows::Forms::DataGridViewColumnHeadersHeightSizeMode::AutoSize;
			this->dataGridView1->Columns->AddRange(gcnew cli::array< System::Windows::Forms::DataGridViewColumn^  >(4) {
				this->id, this->name,
					this->price, this->quantity
			});
			this->dataGridView1->Location = System::Drawing::Point(12, 12);
			this->dataGridView1->Name = L"dataGridView1";
			this->dataGridView1->Size = System::Drawing::Size(443, 290);
			this->dataGridView1->TabIndex = 0;
			// 
			// id
			// 
			this->id->HeaderText = L"ID";
			this->id->Name = L"id";
			// 
			// name
			// 
			this->name->HeaderText = L"Name";
			this->name->Name = L"name";
			// 
			// price
			// 
			this->price->HeaderText = L"Price";
			this->price->Name = L"price";
			// 
			// quantity
			// 
			this->quantity->HeaderText = L"quantity";
			this->quantity->Name = L"quantity";
			// 
			// groupBox1
			// 
			this->groupBox1->Controls->Add(this->button_delete);
			this->groupBox1->Controls->Add(this->button_update);
			this->groupBox1->Controls->Add(this->button_add);
			this->groupBox1->Controls->Add(this->button_download);
			this->groupBox1->Location = System::Drawing::Point(479, 12);
			this->groupBox1->Name = L"groupBox1";
			this->groupBox1->Size = System::Drawing::Size(209, 290);
			this->groupBox1->TabIndex = 1;
			this->groupBox1->TabStop = false;
			this->groupBox1->Text = L"Действия";
			// 
			// button_delete
			// 
			this->button_delete->Location = System::Drawing::Point(41, 223);
			this->button_delete->Name = L"button_delete";
			this->button_delete->Size = System::Drawing::Size(124, 34);
			this->button_delete->TabIndex = 3;
			this->button_delete->Text = L"Удалить";
			this->button_delete->UseVisualStyleBackColor = true;
			this->button_delete->Click += gcnew System::EventHandler(this, &MyForm::button_delete_Click);
			// 
			// button_update
			// 
			this->button_update->Location = System::Drawing::Point(41, 157);
			this->button_update->Name = L"button_update";
			this->button_update->Size = System::Drawing::Size(124, 34);
			this->button_update->TabIndex = 2;
			this->button_update->Text = L"Обновить";
			this->button_update->UseVisualStyleBackColor = true;
			this->button_update->Click += gcnew System::EventHandler(this, &MyForm::button_update_Click);
			// 
			// button_add
			// 
			this->button_add->Location = System::Drawing::Point(41, 94);
			this->button_add->Name = L"button_add";
			this->button_add->Size = System::Drawing::Size(124, 34);
			this->button_add->TabIndex = 1;
			this->button_add->Text = L"Добавить";
			this->button_add->UseVisualStyleBackColor = true;
			this->button_add->Click += gcnew System::EventHandler(this, &MyForm::button_add_Click);
			// 
			// button_download
			// 
			this->button_download->Location = System::Drawing::Point(41, 29);
			this->button_download->Name = L"button_download";
			this->button_download->Size = System::Drawing::Size(124, 34);
			this->button_download->TabIndex = 0;
			this->button_download->Text = L"Загрузить";
			this->button_download->UseVisualStyleBackColor = true;
			this->button_download->Click += gcnew System::EventHandler(this, &MyForm::button_download_Click);
			// 
			// MyForm
			// 
			this->AutoScaleDimensions = System::Drawing::SizeF(6, 13);
			this->AutoScaleMode = System::Windows::Forms::AutoScaleMode::Font;
			this->ClientSize = System::Drawing::Size(703, 314);
			this->Controls->Add(this->groupBox1);
			this->Controls->Add(this->dataGridView1);
			this->Name = L"MyForm";
			this->Text = L"Restaurant Database";
			(cli::safe_cast<System::ComponentModel::ISupportInitialize^>(this->dataGridView1))->EndInit();
			this->groupBox1->ResumeLayout(false);
			this->ResumeLayout(false);

		}
#pragma endregion
private: System::Void button_download_Click(System::Object^ sender, System::EventArgs^ e);
private: System::Void button_add_Click(System::Object^ sender, System::EventArgs^ e);
private: System::Void button_update_Click(System::Object^ sender, System::EventArgs^ e);
private: System::Void button_delete_Click(System::Object^ sender, System::EventArgs^ e);
};
}
