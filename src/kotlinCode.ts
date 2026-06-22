// Kotlin source code files for the production-ready Android Application

export interface KotlinFile {
  name: string;
  path: string;
  language: string;
  content: string;
}

export const kotlinFiles: KotlinFile[] = [
  {
    name: "Deposit.kt",
    path: "app/src/main/java/by/invest/tracker/data/model/Deposit.kt",
    language: "kotlin",
    content: `package by.invest.tracker.data.model

import androidx.room.Entity
import androidx.room.PrimaryKey

@Entity(tableName = "deposits")
data class Deposit(
    @PrimaryKey(autoGenerate = true)
    val id: Long = 0,
    val date: String, // Формат: ГГГГ-ММ-ДД
    val amountByn: Double,
    val amountUsd: Double,
    val commissionByn: Double,
    val commissionUsd: Double
)`
  },
  {
    name: "DepositDao.kt",
    path: "app/src/main/java/by/invest/tracker/data/local/DepositDao.kt",
    language: "kotlin",
    content: `package by.invest.tracker.data.local

import androidx.room.*
import by.invest.tracker.data.model.Deposit
import kotlinx.coroutines.flow.Flow

@Dao
interface DepositDao {
    @Query("SELECT * FROM deposits ORDER BY date DESC, id DESC")
    fun getAllDepositsDesc(): Flow<List<Deposit>>

    @Query("SELECT * FROM deposits ORDER BY date ASC, id ASC")
    fun getAllDepositsAsc(): Flow<List<Deposit>>

    @Insert(onConflict = OnConflictStrategy.REPLACE)
    suspend fun insertDeposit(deposit: Deposit): Long

    @Insert(onConflict = OnConflictStrategy.REPLACE)
    suspend fun insertDeposits(deposits: List<Deposit>)

    @Update
    suspend fun updateDeposit(deposit: Deposit)

    @Delete
    suspend fun deleteDeposit(deposit: Deposit)

    @Query("DELETE FROM deposits")
    suspend fun deleteAll()
}`
  },
  {
    name: "AppDatabase.kt",
    path: "app/src/main/java/by/invest/tracker/data/local/AppDatabase.kt",
    language: "kotlin",
    content: `package by.invest.tracker.data.local

import android.content.Context
import androidx.room.Database
import androidx.room.Room
import androidx.room.RoomDatabase
import by.invest.tracker.data.model.Deposit

@Database(entities = [Deposit::class], version = 1, exportSchema = false)
abstract class AppDatabase : RoomDatabase() {
    abstract fun depositDao(): DepositDao

    companion object {
        @Volatile
        private var INSTANCE: AppDatabase? = null

        fun getDatabase(context: Context): AppDatabase {
            return INSTANCE ?: synchronized(this) {
                val instance = Room.databaseBuilder(
                    context.applicationContext,
                    AppDatabase::class.java,
                    "investment_tracker_db"
                ).fallbackToDestructiveMigration()
                 .build()
                INSTANCE = instance
                instance
            }
        }
    }
}`
  },
  {
    name: "DepositRepository.kt",
    path: "app/src/main/java/by/invest/tracker/data/repository/DepositRepository.kt",
    language: "kotlin",
    content: `package by.invest.tracker.data.repository

import by.invest.tracker.data.model.Deposit
import kotlinx.coroutines.flow.Flow

interface DepositRepository {
    fun getDeposits(isDesc: Boolean): Flow<List<Deposit>>
    suspend fun insertDeposit(deposit: Deposit): Long
    suspend fun insertDeposits(deposits: List<Deposit>)
    suspend fun updateDeposit(deposit: Deposit)
    suspend fun deleteDeposit(deposit: Deposit)
}`
  },
  {
    name: "DepositRepositoryImpl.kt",
    path: "app/src/main/java/by/invest/tracker/data/repository/DepositRepositoryImpl.kt",
    language: "kotlin",
    content: `package by.invest.tracker.data.repository

import by.invest.tracker.data.local.DepositDao
import by.invest.tracker.data.model.Deposit
import kotlinx.coroutines.flow.Flow

class DepositRepositoryImpl(private val depositDao: DepositDao) : DepositRepository {
    override fun getDeposits(isDesc: Boolean): Flow<List<Deposit>> {
        return if (isDesc) {
            depositDao.getAllDepositsDesc()
        } else {
            depositDao.getAllDepositsAsc()
        }
    }

    override suspend fun insertDeposit(deposit: Deposit): Long {
        return depositDao.insertDeposit(deposit)
    }

    override suspend fun insertDeposits(deposits: List<Deposit>) {
        depositDao.insertDeposits(deposits)
    }

    override suspend fun updateDeposit(deposit: Deposit) {
        depositDao.updateDeposit(deposit)
    }

    override suspend fun deleteDeposit(deposit: Deposit) {
        depositDao.deleteDeposit(deposit)
    }
}`
  },
  {
    name: "ExcelUtils.kt",
    path: "app/src/main/java/by/invest/tracker/util/ExcelUtils.kt",
    language: "kotlin",
    content: `package by.invest.tracker.util

import android.content.Context
import android.net.Uri
import by.invest.tracker.data.model.Deposit
import org.apache.poi.xssf.usermodel.XSSFWorkbook
import java.io.InputStream
import java.io.OutputStream

object ExcelUtils {

    /**
     * Экспорт депозитов в поток вывода XLSX (через SAF Uri)
     */
    fun exportToExcel(outputStream: OutputStream, deposits: List<Deposit>): Boolean {
        return try {
            val workbook = XSSFWorkbook()
            val sheet = workbook.createSheet("Депозиты")

            // Заголовок
            val headerRow = sheet.createRow(0)
            val columns = arrayOf("ID", "Дата", "Сумма BYN", "Сумма USD", "Комиссия BYN", "Комиссия USD")
            
            val headerStyle = workbook.createCellStyle().apply {
                val font = workbook.createFont().apply {
                    bold = true
                }
                setFont(font)
            }

            for (i in columns.indices) {
                val cell = headerRow.createCell(i)
                cell.setCellValue(columns[i])
                cell.cellStyle = headerStyle
            }

            // Заполнение данными
            for (rowIndex in deposits.indices) {
                val deposit = deposits[rowIndex]
                val row = sheet.createRow(rowIndex + 1)
                
                row.createCell(0).setCellValue(deposit.id.toDouble())
                row.createCell(1).setCellValue(deposit.date)
                row.createCell(2).setCellValue(deposit.amountByn)
                row.createCell(3).setCellValue(deposit.amountUsd)
                row.createCell(4).setCellValue(deposit.commissionByn)
                row.createCell(5).setCellValue(deposit.commissionUsd)
            }

            // Авторазмер колонок
            for (i in columns.indices) {
                sheet.autoSizeColumn(i)
            }

            workbook.write(outputStream)
            workbook.close()
            outputStream.close()
            true
        } catch (e: Exception) {
            e.printStackTrace()
            false
        }
    }

    /**
     * Импорт списка депозитов из входного потока XLSX
     */
    fun importFromExcel(inputStream: InputStream): List<Deposit>? {
        return try {
            val deposits = mutableListOf<Deposit>()
            val workbook = XSSFWorkbook(inputStream)
            val sheet = workbook.getSheetAt(0) ?: return null

            val rowIterator = sheet.iterator()
            if (!rowIterator.hasNext()) {
                workbook.close()
                return null
            }

            // Пропускаем шапку
            rowIterator.next()

            while (rowIterator.hasNext()) {
                val row = rowIterator.next()
                
                // Проверяем, что дата присутствует
                val dateCell = row.getCell(1) ?: continue
                val dateStr = when (dateCell.cellType) {
                    org.apache.poi.ss.usermodel.CellType.STRING -> dateCell.stringCellValue
                    else -> dateCell.toString()
                }
                if (dateStr.isBlank()) continue

                // Считываем значения ячеек безопасно
                val amountByn = row.getCell(2)?.numericCellValue ?: 0.0
                val amountUsd = row.getCell(3)?.numericCellValue ?: 0.0
                val commissionByn = row.getCell(4)?.numericCellValue ?: 0.0
                val commissionUsd = row.getCell(5)?.numericCellValue ?: 0.0

                deposits.add(
                    Deposit(
                        date = dateStr.trim(),
                        amountByn = amountByn,
                        amountUsd = amountUsd,
                        commissionByn = commissionByn,
                        commissionUsd = commissionUsd
                    )
                )
            }

            workbook.close()
            inputStream.close()
            deposits
        } catch (e: Exception) {
            e.printStackTrace()
            null
        }
    }
}`
  },
  {
    name: "DepositViewModel.kt",
    path: "app/src/main/java/by/invest/tracker/ui/viewmodel/DepositViewModel.kt",
    language: "kotlin",
    content: `package by.invest.tracker.ui.viewmodel

import android.app.Application
import android.content.Context
import android.net.Uri
import androidx.lifecycle.AndroidViewModel
import androidx.lifecycle.viewModelScope
import by.invest.tracker.data.local.AppDatabase
import by.invest.tracker.data.model.Deposit
import by.invest.tracker.data.repository.DepositRepositoryImpl
import by.invest.tracker.util.ExcelUtils
import kotlinx.coroutines.ExperimentalCoroutinesApi
import kotlinx.coroutines.flow.*
import kotlinx.coroutines.launch
import java.text.SimpleDateFormat
import java.util.*

class DepositViewModel(application: Application) : AndroidViewModel(application) {

    private val repository: DepositRepositoryImpl = DepositRepositoryImpl(AppDatabase.getDatabase(application).depositDao())
    
    // Текущий режим сортировки: true - новые сначала, false - старые сначала
    private val _isSortDesc = MutableStateFlow(true)
    val isSortDesc: StateFlow<Boolean> = _isSortDesc.asStateFlow()

    // Наблюдение за всеми записями на основе сортировки
    @OptIn(ExperimentalCoroutinesApi::class)
    val deposits: StateFlow<List<Deposit>> = _isSortDesc
        .flatMapLatest { desc ->
            repository.getDeposits(desc)
        }
        .stateIn(viewModelScope, SharingStarted.WhileSubscribed(5000), emptyList())

    // Состояние пагинации
    private val _currentPage = MutableStateFlow(1)
    val currentPage: StateFlow<Int> = _currentPage.asStateFlow()
    val pageSize = 12

    // Текущий список депозитов на странице
    val pagedDeposits: StateFlow<List<Deposit>> = combine(deposits, _currentPage) { list, page ->
        val startIndex = (page - 1) * pageSize
        if (startIndex >= list.size) {
            _currentPage.value = 1
            list.take(pageSize)
        } else {
            list.drop(startIndex).take(pageSize)
        }
    }.stateIn(viewModelScope, SharingStarted.WhileSubscribed(5000), emptyList())

    // Суммарные метрики
    val totalInvestedByn: StateFlow<Double> = deposits.map { list -> list.sumOf { it.amountByn } }
        .stateIn(viewModelScope, SharingStarted.WhileSubscribed(5000), 0.0)

    val totalInvestedUsd: StateFlow<Double> = deposits.map { list -> list.sumOf { it.amountUsd } }
        .stateIn(viewModelScope, SharingStarted.WhileSubscribed(5000), 0.0)

    val totalCommissionByn: StateFlow<Double> = deposits.map { list -> list.sumOf { it.commissionByn } }
        .stateIn(viewModelScope, SharingStarted.WhileSubscribed(5000), 0.0)

    val totalCommissionUsd: StateFlow<Double> = deposits.map { list -> list.sumOf { it.commissionUsd } }
        .stateIn(viewModelScope, SharingStarted.WhileSubscribed(5000), 0.0)

    // Буфер для реализации Undo (Отмена действия)
    private var lastDeletedDeposit: Deposit? = null
    private var lastEditedDepositBeforeEdit: Deposit? = null
    
    private val _eventFlow = MutableSharedFlow<UiEvent>()
    val eventFlow = _eventFlow.asSharedFlow()

    init {
    }

    fun toggleSorting() {
        _isSortDesc.value = !_isSortDesc.value
        _currentPage.value = 1
    }

    fun setPage(page: Int) {
        _currentPage.value = page
    }

    fun addDeposit(date: String, amountByn: Double, amountUsd: Double, commissionByn: Double, commissionUsd: Double) {
        viewModelScope.launch {
            val deposit = Deposit(
                date = date,
                amountByn = amountByn,
                amountUsd = amountUsd,
                commissionByn = commissionByn,
                commissionUsd = commissionUsd
            )
            repository.insertDeposit(deposit)
            _eventFlow.emit(UiEvent.ShowSnackbar("Депозит добавлен успешно!"))
        }
    }

    fun updateDeposit(deposit: Deposit) {
        viewModelScope.launch {
            // Сохраняем предыдущую копию для отката
            val original = deposits.value.find { it.id == deposit.id }
            if (original != null) {
                lastEditedDepositBeforeEdit = original
            }
            repository.updateDeposit(deposit)
            _eventFlow.emit(UiEvent.ShowSnackbar("Запись обновлена!", showUndoEdit = true))
        }
    }

    fun deleteDeposit(deposit: Deposit) {
        viewModelScope.launch {
            lastDeletedDeposit = deposit
            repository.deleteDeposit(deposit)
            _eventFlow.emit(UiEvent.ShowSnackbar("Запись удалена!", showUndoDelete = true))
        }
    }

    fun undoDelete() {
        val depositToRestore = lastDeletedDeposit ?: return
        viewModelScope.launch {
            repository.insertDeposit(depositToRestore)
            lastDeletedDeposit = null
            _eventFlow.emit(UiEvent.ShowSnackbar("Удаление отменено!"))
        }
    }

    fun undoEdit() {
        val depositToRestore = lastEditedDepositBeforeEdit ?: return
        viewModelScope.launch {
            repository.updateDeposit(depositToRestore)
            lastEditedDepositBeforeEdit = null
            _eventFlow.emit(UiEvent.ShowSnackbar("Изменение отменено!"))
        }
    }

    fun exportData(context: Context, uri: Uri) {
        viewModelScope.launch {
            try {
                val outputStream = context.contentResolver.openOutputStream(uri)
                if (outputStream != null) {
                    val success = ExcelUtils.exportToExcel(outputStream, deposits.value)
                    if (success) {
                        _eventFlow.emit(UiEvent.ShowSnackbar("Данные экспортированы в XLSX!"))
                    } else {
                        _eventFlow.emit(UiEvent.ShowSnackbar("Ошибка экспорта данных!"))
                    }
                }
            } catch (e: Exception) {
                _eventFlow.emit(UiEvent.ShowSnackbar("Ошибка: \${e.message}"))
            }
        }
    }

    fun importData(context: Context, uri: Uri) {
        viewModelScope.launch {
            try {
                val inputStream = context.contentResolver.openInputStream(uri)
                if (inputStream != null) {
                    val importedList = ExcelUtils.importFromExcel(inputStream)
                    if (importedList != null) {
                        repository.insertDeposits(importedList)
                        _eventFlow.emit(UiEvent.ShowSnackbar("Импортировано \${importedList.size} записей!"))
                    } else {
                        _eventFlow.emit(UiEvent.ShowSnackbar("Не удалось прочитать Excel файл!"))
                    }
                }
            } catch (e: Exception) {
                _eventFlow.emit(UiEvent.ShowSnackbar("Ошибка импорта: \${e.message}"))
            }
        }
    }

    sealed class UiEvent {
        data class ShowSnackbar(
            val message: String, 
            val showUndoDelete: Boolean = false,
            val showUndoEdit: Boolean = false
        ) : UiEvent()
    }
}`
  },
  {
    name: "MainActivity.kt",
    path: "app/src/main/java/by/invest/tracker/MainActivity.kt",
    language: "kotlin",
    content: `package by.invest.tracker

import android.os.Bundle
import androidx.activity.ComponentActivity
import androidx.activity.compose.setContent
import androidx.lifecycle.viewmodel.compose.viewModel
import androidx.compose.foundation.layout.fillMaxSize
import androidx.compose.material3.MaterialTheme
import androidx.compose.material3.Surface
import androidx.compose.ui.Modifier
import by.invest.tracker.ui.screen.DepositScreen
import by.invest.tracker.ui.theme.InvestmentTrackerTheme
import by.invest.tracker.ui.viewmodel.DepositViewModel

class MainActivity : ComponentActivity() {
    override fun onCreate(savedInstanceState: Bundle?) {
        super.onCreate(savedInstanceState)
        setContent {
            InvestmentTrackerTheme {
                Surface(
                    modifier = Modifier.fillMaxSize(),
                    color = MaterialTheme.colorScheme.background
                ) {
                    val viewModel: DepositViewModel = viewModel()
                    DepositScreen(viewModel = viewModel)
                }
            }
        }
    }
}`
  },
  {
    name: "DepositScreen.kt",
    path: "app/src/main/java/by/invest/tracker/ui/screen/DepositScreen.kt",
    language: "kotlin",
    content: `package by.invest.tracker.ui.screen

import android.app.DatePickerDialog
import androidx.activity.compose.rememberLauncherForActivityResult
import androidx.activity.result.contract.ActivityResultContracts
import androidx.compose.animation.*
import androidx.compose.foundation.background
import androidx.compose.foundation.clickable
import androidx.compose.foundation.layout.*
import androidx.compose.foundation.lazy.LazyColumn
import androidx.compose.foundation.lazy.items
import androidx.compose.foundation.shape.RoundedCornerShape
import androidx.compose.foundation.text.KeyboardOptions
import androidx.compose.material.icons.Icons
import androidx.compose.material.icons.filled.*
import androidx.compose.material3.*
import androidx.compose.runtime.*
import androidx.compose.ui.Alignment
import androidx.compose.ui.Modifier
import androidx.compose.ui.draw.clip
import androidx.compose.ui.platform.LocalContext
import androidx.compose.ui.text.font.FontWeight
import androidx.compose.ui.text.input.KeyboardType
import androidx.compose.ui.text.style.TextAlign
import androidx.compose.ui.unit.dp
import androidx.compose.ui.unit.sp
import by.invest.tracker.data.model.Deposit
import by.invest.tracker.ui.viewmodel.DepositViewModel
import kotlinx.coroutines.flow.collectLatest
import java.text.SimpleDateFormat
import java.util.*

@OptIn(ExperimentalLayoutApi::class, ExperimentalMaterial3Api::class)
@Composable
fun DepositScreen(viewModel: DepositViewModel) {
    val context = LocalContext.current
    val scope = rememberCoroutineScope()
    val snackbarHostState = remember { SnackbarHostState() }

    // Наблюдение за состояниями из ViewModel
    val pagedDeposits by viewModel.pagedDeposits.collectAsState()
    val totalByn by viewModel.totalInvestedByn.collectAsState()
    val totalUsd by viewModel.totalInvestedUsd.collectAsState()
    val commissionsByn by viewModel.totalCommissionByn.collectAsState()
    val commissionsUsd by viewModel.totalCommissionUsd.collectAsState()
    val isSortDesc by viewModel.isSortDesc.collectAsState()
    val currentPage by viewModel.currentPage.collectAsState()
    val allDeposits by viewModel.deposits.collectAsState()

    // Рассчитываем количество страниц
    val totalPages = remember(allDeposits) {
        val pages = (allDeposits.size + viewModel.pageSize - 1) / viewModel.pageSize
        if (pages < 1) 1 else pages
    }

    // Состояния полей ввода
    var amountBynInput by remember { mutableStateOf("") }
    var amountUsdInput by remember { mutableStateOf("") }
    var commissionBynInput by remember { mutableStateOf("") }
    var commissionUsdInput by remember { mutableStateOf("") }

    // Умная интуитивная дата: предзаполняется сегодняшней датой
    val russiandDateFormat = remember { SimpleDateFormat("yyyy-MM-dd", Locale.getDefault()) }
    var selectedDate by remember { mutableStateOf(russiandDateFormat.format(Date())) }

    // Диалог редактирования
    var editingDeposit by remember { mutableStateOf<Deposit?>(null) }

    // Слушатель событий из ViewModel (для снекбаров с Undo)
    LaunchedEffect(key1 = true) {
        viewModel.eventFlow.collectLatest { event ->
            when (event) {
                is DepositViewModel.UiEvent.ShowSnackbar -> {
                    val result = snackbarHostState.showSnackbar(
                        message = event.message,
                        actionLabel = if (event.showUndoDelete || event.showUndoEdit) "Отменить" else null,
                        duration = SnackbarDuration.Short
                    )
                    if (result == SnackbarResult.ActionPerformed) {
                        if (event.showUndoDelete) {
                            viewModel.undoDelete()
                        } else if (event.showUndoEdit) {
                            viewModel.undoEdit()
                        }
                    }
                }
            }
        }
    }

    // SAF Лаунчеры для экспорта и импорта Excel файлов
    val exportLauncher = rememberLauncherForActivityResult(
        contract = ActivityResultContracts.CreateDocument("application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"),
        onResult = { uri ->
            uri?.let { viewModel.exportData(context, it) }
        }
    )

    val importLauncher = rememberLauncherForActivityResult(
        contract = ActivityResultContracts.OpenDocument(),
        onResult = { uri ->
            uri?.let { viewModel.importData(context, it) }
        }
    )

    // Функция выбора даты
    val calendar = Calendar.getInstance()
    val datePickerDialog = DatePickerDialog(
        context,
        { _, year, month, dayOfMonth ->
            val tempCal = Calendar.getInstance()
            tempCal.set(year, month, dayOfMonth)
            selectedDate = russiandDateFormat.format(tempCal.time)
        },
        calendar.get(Calendar.YEAR),
        calendar.get(Calendar.MONTH),
        calendar.get(Calendar.DAY_OF_MONTH)
    )

    Scaffold(
        snackbarHost = { SnackbarHost(hostState = snackbarHostState) },
        topBar = {
            SmallTopAppBar(
                title = { Text("Учет Депозитов и Сбережений", fontWeight = FontWeight.Bold) },
                actions = {
                    IconButton(onClick = { importLauncher.launch(arrayOf("application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")) }) {
                        Icon(Icons.Filled.ArrowUpward, contentDescription = "Импорт XLSX")
                    }
                    IconButton(onClick = { exportLauncher.launch("депозиты_\${System.currentTimeMillis()}.xlsx") }) {
                        Icon(Icons.Filled.ArrowDownward, contentDescription = "Экспорт XLSX")
                    }
                },
                colors = TopAppBarDefaults.smallTopAppBarColors(
                    containerColor = MaterialTheme.colorScheme.primaryContainer,
                    titleContentColor = MaterialTheme.colorScheme.onPrimaryContainer
                )
            )
        }
    ) { paddingValues ->
        LazyColumn(
            modifier = Modifier
                .fillMaxSize()
                .padding(paddingValues)
                .padding(16.dp),
            verticalArrangement = Arrangement.spacedBy(16.dp)
        ) {
            // 1. ШАПКА МЕТРИК
            item {
                Card(
                    colors = CardDefaults.cardColors(
                        containerColor = MaterialTheme.colorScheme.secondaryContainer
                    ),
                    modifier = Modifier.fillMaxWidth()
                ) {
                    Column(
                        modifier = Modifier.padding(16.dp),
                        verticalArrangement = Arrangement.spacedBy(8.dp)
                    ) {
                        Text(
                            text = "Финансовые индикаторы",
                            fontSize = 18.sp,
                            fontWeight = FontWeight.Bold,
                            color = MaterialTheme.colorScheme.onSecondaryContainer
                        )
                        Divider(color = MaterialTheme.colorScheme.onSecondaryContainer.copy(alpha = 0.2f))
                        
                        Row(
                            modifier = Modifier.fillMaxWidth(),
                            horizontalArrangement = Arrangement.SpaceBetween
                        ) {
                            Column(modifier = Modifier.weight(1f)) {
                                Text("Вложено BYN:", fontSize = 12.sp, color = MaterialTheme.colorScheme.onSecondaryContainer.copy(alpha = 0.7f))
                                Text(String.format(Locale.US, "%,.2f BYN", totalByn), fontSize = 16.sp, fontWeight = FontWeight.Bold)
                            }
                            Column(modifier = Modifier.weight(1f)) {
                                Text("Вложено USD:", fontSize = 12.sp, color = MaterialTheme.colorScheme.onSecondaryContainer.copy(alpha = 0.7f))
                                Text(String.format(Locale.US, "$%,.2f", totalUsd), fontSize = 16.sp, fontWeight = FontWeight.Bold, color = MaterialTheme.colorScheme.primary)
                            }
                        }

                        Row(
                            modifier = Modifier.fillMaxWidth(),
                            horizontalArrangement = Arrangement.SpaceBetween
                        ) {
                            Column(modifier = Modifier.weight(1f)) {
                                Text("Комиссии BYN:", fontSize = 12.sp, color = MaterialTheme.colorScheme.onSecondaryContainer.copy(alpha = 0.7f))
                                Text(String.format(Locale.US, "%,.2f BYN", commissionsByn), fontSize = 14.sp, fontWeight = FontWeight.SemiBold)
                            }
                            Column(modifier = Modifier.weight(1f)) {
                                Text("Комиссии USD:", fontSize = 12.sp, color = MaterialTheme.colorScheme.onSecondaryContainer.copy(alpha = 0.7f))
                                Text(String.format(Locale.US, "$%,.2f", commissionsUsd), fontSize = 14.sp, fontWeight = FontWeight.SemiBold)
                            }
                        }
                    }
                }
            }

            // 2. ФОРМА ВВОДА
            item {
                Card(
                    modifier = Modifier.fillMaxWidth()
                ) {
                    Column(
                        modifier = Modifier.padding(16.dp),
                        verticalArrangement = Arrangement.spacedBy(12.dp)
                    ) {
                        Text("Добавить новый депозит", fontWeight = FontWeight.Bold, fontSize = 16.sp)

                        Row(
                            modifier = Modifier.fillMaxWidth(),
                            horizontalArrangement = Arrangement.spacedBy(8.dp)
                        ) {
                            OutlinedTextField(
                                value = amountBynInput,
                                onValueChange = { amountBynInput = it },
                                label = { Text("Сумма BYN") },
                                keyboardOptions = KeyboardOptions(keyboardType = KeyboardType.Number),
                                modifier = Modifier.weight(1f)
                            )
                            OutlinedTextField(
                                value = amountUsdInput,
                                onValueChange = { amountUsdInput = it },
                                label = { Text("Сумма USD") },
                                keyboardOptions = KeyboardOptions(keyboardType = KeyboardType.Number),
                                modifier = Modifier.weight(1f)
                            )
                        }

                        Row(
                            modifier = Modifier.fillMaxWidth(),
                            horizontalArrangement = Arrangement.spacedBy(8.dp)
                        ) {
                            OutlinedTextField(
                                value = commissionBynInput,
                                onValueChange = { commissionBynInput = it },
                                label = { Text("Комиссия BYN") },
                                keyboardOptions = KeyboardOptions(keyboardType = KeyboardType.Number),
                                modifier = Modifier.weight(1f)
                            )
                            OutlinedTextField(
                                value = commissionUsdInput,
                                onValueChange = { commissionUsdInput = it },
                                label = { Text("Комиссия USD") },
                                keyboardOptions = KeyboardOptions(keyboardType = KeyboardType.Number),
                                modifier = Modifier.weight(1f)
                            )
                        }

                        // Умная дата
                        Row(
                            modifier = Modifier
                                .fillMaxWidth()
                                .clip(RoundedCornerShape(8.dp))
                                .background(MaterialTheme.colorScheme.surfaceVariant)
                                .clickable { datePickerDialog.show() }
                                .padding(12.dp),
                            verticalAlignment = Alignment.CenterVertically,
                            horizontalArrangement = Arrangement.SpaceBetween
                        ) {
                            Row(verticalAlignment = Alignment.CenterVertically) {
                                Icon(Icons.Filled.DateRange, contentDescription = "Календарь", tint = MaterialTheme.colorScheme.primary)
                                Spacer(modifier = Modifier.width(8.dp))
                                Text("Дата транзакции:", fontWeight = FontWeight.Medium)
                            }
                            Text(
                                text = selectedDate,
                                fontWeight = FontWeight.Bold,
                                color = MaterialTheme.colorScheme.primary,
                                fontSize = 16.sp
                            )
                        }

                        Row(
                            modifier = Modifier.fillMaxWidth(),
                            horizontalArrangement = Arrangement.spacedBy(8.dp)
                        ) {
                            Button(
                                onClick = {
                                    val bAmount = amountBynInput.toDoubleOrNull() ?: 0.0
                                    val uAmount = amountUsdInput.toDoubleOrNull() ?: 0.0
                                    val bComm = commissionBynInput.toDoubleOrNull() ?: 0.0
                                    val uComm = commissionUsdInput.toDoubleOrNull() ?: 0.0

                                    if (bAmount > 0 || uAmount > 0) {
                                        viewModel.addDeposit(selectedDate, bAmount, uAmount, bComm, uComm)
                                        // Очистка
                                        amountBynInput = ""
                                        amountUsdInput = ""
                                        commissionBynInput = ""
                                        commissionUsdInput = ""
                                        // Автоматически возвращается к сегодняшней
                                        selectedDate = russiandDateFormat.format(Date())
                                    }
                                },
                                modifier = Modifier.weight(1.5f)
                            ) {
                                Icon(Icons.Filled.Add, contentDescription = null)
                                Spacer(modifier = Modifier.width(4.dp))
                                Text("Депозит")
                            }

                            FilledTonalButton(
                                onClick = { viewModel.toggleSorting() },
                                modifier = Modifier.weight(1f)
                            ) {
                                Icon(
                                    imageVector = if (isSortDesc) Icons.Filled.SortByAlpha else Icons.Filled.Sort,
                                    contentDescription = null
                                )
                                Spacer(modifier = Modifier.width(4.dp))
                                Text(
                                    text = if (isSortDesc) "Сначала новые" else "Сначала старые",
                                    fontSize = 11.sp,
                                    textAlign = TextAlign.Center
                                )
                            }
                        }
                    }
                }
            }

            // ЗАГОЛОВОК СПИСКА
            item {
                Row(
                    modifier = Modifier.fillMaxWidth(),
                    horizontalArrangement = Arrangement.SpaceBetween,
                    verticalAlignment = Alignment.CenterVertically
                ) {
                    Text(
                        text = "История депозитов (\${allDeposits.size})",
                        fontWeight = FontWeight.Bold,
                        fontSize = 18.sp
                    )
                    Text(
                        text = "Стр. \$currentPage из \$totalPages",
                        fontSize = 12.sp,
                        color = MaterialTheme.colorScheme.onSurfaceVariant
                    )
                }
            }

            // 3. СПИСОК С ПАГИНАЦИЕЙ (РЕКОРДЫ)
            if (pagedDeposits.isEmpty()) {
                item {
                    Box(
                        modifier = Modifier
                            .fillMaxWidth()
                            .padding(vertical = 32.dp),
                        contentAlignment = Alignment.Center
                    ) {
                        Text("Нет записей об инвестициях", color = MaterialTheme.colorScheme.onSurfaceVariant)
                    }
                }
            } else {
                items(pagedDeposits, key = { it.id }) { deposit ->
                    Card(
                        colors = CardDefaults.cardColors(
                            containerColor = MaterialTheme.colorScheme.surface
                        ),
                        elevation = CardDefaults.cardElevation(2.dp),
                        modifier = Modifier.fillMaxWidth()
                    ) {
                        Row(
                            modifier = Modifier
                                .fillMaxWidth()
                                .padding(12.dp),
                            verticalAlignment = Alignment.CenterVertically,
                            horizontalArrangement = Arrangement.SpaceBetween
                        ) {
                            Column(modifier = Modifier.weight(1f)) {
                                Text(
                                    text = deposit.date,
                                    fontWeight = FontWeight.Bold,
                                    color = MaterialTheme.colorScheme.primary,
                                    fontSize = 14.sp
                                )
                                Spacer(modifier = Modifier.height(4.dp))
                                Row(
                                    horizontalArrangement = Arrangement.spacedBy(16.dp),
                                    modifier = Modifier.fillMaxWidth()
                                ) {
                                    Column {
                                        Text("Сумма:", fontSize = 11.sp, color = MaterialTheme.colorScheme.onSurfaceVariant)
                                        Text("\${deposit.amountByn} BYN", fontSize = 13.sp, fontWeight = FontWeight.Medium)
                                        Text("\$ \${deposit.amountUsd}", fontSize = 13.sp, fontWeight = FontWeight.Medium, color = MaterialTheme.colorScheme.primary)
                                    }
                                    Column {
                                        Text("Комиссия:", fontSize = 11.sp, color = MaterialTheme.colorScheme.onSurfaceVariant)
                                        Text("\${deposit.commissionByn} BYN", fontSize = 12.sp, color = MaterialTheme.colorScheme.outline)
                                        Text("\$ \${deposit.commissionUsd}", fontSize = 12.sp, color = MaterialTheme.colorScheme.outline)
                                    }
                                }
                            }

                            Row {
                                IconButton(onClick = { editingDeposit = deposit }) {
                                    Icon(Icons.Filled.Edit, contentDescription = "Редактировать", tint = MaterialTheme.colorScheme.primary)
                                }
                                IconButton(onClick = { viewModel.deleteDeposit(deposit) }) {
                                    Icon(Icons.Filled.Delete, contentDescription = "Улить", tint = MaterialTheme.colorScheme.error)
                                }
                            }
                        }
                    }
                }
            }

            // 4. КОНТРОЛЛЕР ПАГИНАЦИИ
            if (totalPages > 1) {
                item {
                    FlowRow(
                        modifier = Modifier
                            .fillMaxWidth()
                            .padding(vertical = 8.dp),
                        horizontalArrangement = Arrangement.Center
                    ) {
                        for (page in 1..totalPages) {
                            val isSelected = page == currentPage
                            Button(
                                onClick = { viewModel.setPage(page) },
                                colors = ButtonDefaults.buttonColors(
                                    containerColor = if (isSelected) MaterialTheme.colorScheme.primary else MaterialTheme.colorScheme.surfaceVariant,
                                    contentColor = if (isSelected) MaterialTheme.colorScheme.onPrimary else MaterialTheme.colorScheme.onSurfaceVariant
                                ),
                                contentPadding = PaddingValues(horizontal = 12.dp, vertical = 4.dp),
                                modifier = Modifier
                                    .padding(4.dp)
                                    .sizeIn(minWidth = 36.dp, minHeight = 36.dp)
                            ) {
                                Text(text = page.toString(), fontWeight = FontWeight.Bold)
                            }
                        }
                    }
                }
            }
        }
    }

    // ДИАЛОГ РЕДАКТИРОВАНИЯ
    editingDeposit?.let { deposit ->
        var editByn by remember { mutableStateOf(deposit.amountByn.toString()) }
        var editUsd by remember { mutableStateOf(deposit.amountUsd.toString()) }
        var editCommByn by remember { mutableStateOf(deposit.commissionByn.toString()) }
        var editCommUsd by remember { mutableStateOf(deposit.commissionUsd.toString()) }
        var editDate by remember { mutableStateOf(deposit.date) }

        val editDatePicker = DatePickerDialog(
            context,
            { _, year, month, dayOfMonth ->
                val tempCal = Calendar.getInstance()
                tempCal.set(year, month, dayOfMonth)
                editDate = russiandDateFormat.format(tempCal.time)
            },
            calendar.get(Calendar.YEAR),
            calendar.get(Calendar.MONTH),
            calendar.get(Calendar.DAY_OF_MONTH)
        )

        AlertDialog(
            onDismissRequest = { editingDeposit = null },
            title = { Text("Редактировать депозит", fontWeight = FontWeight.Bold) },
            text = {
                Column(verticalArrangement = Arrangement.spacedBy(10.dp)) {
                    Text("Измените необходимые параметры транзакции:")
                    
                    Row(horizontalArrangement = Arrangement.spacedBy(8.dp)) {
                        OutlinedTextField(
                            value = editByn,
                            onValueChange = { editByn = it },
                            label = { Text("Сумма BYN") },
                            keyboardOptions = KeyboardOptions(keyboardType = KeyboardType.Number),
                            modifier = Modifier.weight(1f)
                        )
                        OutlinedTextField(
                            value = editUsd,
                            onValueChange = { editUsd = it },
                            label = { Text("Сумма USD") },
                            keyboardOptions = KeyboardOptions(keyboardType = KeyboardType.Number),
                            modifier = Modifier.weight(1f)
                        )
                    }

                    Row(horizontalArrangement = Arrangement.spacedBy(8.dp)) {
                        OutlinedTextField(
                            value = editCommByn,
                            onValueChange = { editCommByn = it },
                            label = { Text("Комиссия BYN") },
                            keyboardOptions = KeyboardOptions(keyboardType = KeyboardType.Number),
                            modifier = Modifier.weight(1f)
                        )
                        OutlinedTextField(
                            value = editCommUsd,
                            onValueChange = { editCommUsd = it },
                            label = { Text("Комиссия USD") },
                            keyboardOptions = KeyboardOptions(keyboardType = KeyboardType.Number),
                            modifier = Modifier.weight(1f)
                        )
                    }

                    Row(
                        modifier = Modifier
                            .fillMaxWidth()
                            .clip(RoundedCornerShape(8.dp))
                            .background(MaterialTheme.colorScheme.surfaceVariant)
                            .clickable { editDatePicker.show() }
                            .padding(10.dp),
                        verticalAlignment = Alignment.CenterVertically,
                        horizontalArrangement = Arrangement.SpaceBetween
                    ) {
                        Row(verticalAlignment = Alignment.CenterVertically) {
                            Icon(Icons.Filled.DateRange, contentDescription = null, tint = MaterialTheme.colorScheme.primary)
                            Spacer(modifier = Modifier.width(6.dp))
                            Text("Дата:")
                        }
                        Text(editDate, fontWeight = FontWeight.Bold, color = MaterialTheme.colorScheme.primary)
                    }
                }
            },
            confirmButton = {
                Button(
                    onClick = {
                        val updated = deposit.copy(
                            date = editDate,
                            amountByn = editByn.toDoubleOrNull() ?: 0.0,
                            amountUsd = editUsd.toDoubleOrNull() ?: 0.0,
                            commissionByn = editCommByn.toDoubleOrNull() ?: 0.0,
                            commissionUsd = editCommUsd.toDoubleOrNull() ?: 0.0
                        )
                        viewModel.updateDeposit(updated)
                        editingDeposit = null
                    }
                ) {
                    Text("Сохранить")
                }
            },
            dismissButton = {
                TextButton(onClick = { editingDeposit = null }) {
                    Text("Cancel")
                }
            }
        )
    }
}
`
  },
  {
    name: "app/build.gradle.kts",
    path: "app/build.gradle.kts",
    language: "kotlin",
    content: `plugins {
    id("com.android.application")
    id("org.jetbrains.kotlin.android")
    id("org.jetbrains.kotlin.kapt")
}

android {
    namespace = "by.invest.tracker"
    compileSdk = 34

    defaultConfig {
        applicationId = "by.invest.tracker"
        minSdk = 26
        targetSdk = 34
        versionCode = 1
        versionName = "1.0.0"

        testInstrumentationRunner = "androidx.test.runner.AndroidJUnitRunner"
        vectorDrawables {
            useSupportLibrary = true
        }
    }

    buildTypes {
        release {
            isMinifyEnabled = false
            proguardFiles(
                getDefaultProguardFile("proguard-android-optimize.txt"),
                "proguard-rules.pro"
            )
        }
    }
    compileOptions {
        sourceCompatibility = JavaVersion.VERSION_18
        targetCompatibility = JavaVersion.VERSION_18
    }
    kotlinOptions {
        jvmTarget = "18"
    }
    buildFeatures {
        compose = true
    }
    composeOptions {
        kotlinCompilerExtensionVersion = "1.5.14"
    }
    packaging {
        resources {
            excludes += "/META-INF/{AL2.0,LGPL2.1}"
            // Исключить дубликаты POI лицензий при сборке
            excludes += "META-INF/LICENSE"
            excludes += "META-INF/NOTICE"
            excludes += "META-INF/NOTICE.txt"
            excludes += "META-INF/LICENSE.txt"
        }
    }
}

dependencies {
    // Jetpack Compose
    implementation(platform("androidx.compose:compose-bom:2023.08.00"))
    implementation("androidx.compose.ui:ui")
    implementation("androidx.compose.ui:ui-graphics")
    implementation("androidx.compose.ui:ui-tooling-preview")
    implementation("androidx.compose.material3:material3")
    implementation("androidx.compose.material:material-icons-extended")
    implementation("androidx.activity:activity-compose:1.8.2")
    implementation("androidx.lifecycle:lifecycle-viewmodel-compose:2.7.0")

    // Room DB
    val roomVersion = "2.6.1"
    implementation("androidx.room:room-runtime:$roomVersion")
    implementation("androidx.room:room-ktx:$roomVersion")
    kapt("androidx.room:room-compiler:$roomVersion")

    // Apache POI для импорта / экспорта XLSX
    implementation("org.apache.poi:poi-ooxml:5.2.5") {
        exclude(group = "org.apache.xmlgraphics", module = "batik-all")
        exclude(group = "xml-apis", module = "xml-apis")
    }

    // Coroutines
    implementation("org.jetbrains.kotlinx:kotlinx-coroutines-android:1.7.3")

    // Тестирование
    testImplementation("junit:junit:4.13.2")
    androidTestImplementation("androidx.test.ext:junit:1.1.5")
    androidTestImplementation("androidx.test.espresso:espresso-core:3.5.1")
    androidTestImplementation(platform("androidx.compose:compose-bom:2023.08.00"))
    androidTestImplementation("androidx.compose.ui:ui-test-junit4")
    debugImplementation("androidx.compose.ui:ui-tooling")
    debugImplementation("androidx.compose.ui:ui-test-manifest")
}`
  },
  {
    name: "build.gradle.kts (Root level)",
    path: "build.gradle.kts",
    language: "kotlin",
    content: `// Top-level build file where you can add configuration options common to all sub-projects/modules.
plugins {
    id("com.android.application") version "8.2.2" apply false
    id("org.jetbrains.kotlin.android") version "1.9.24" apply false
}

tasks.register("clean", Delete::class) {
    delete(rootProject.layout.buildDirectory)
}`
  },
  {
    name: "settings.gradle.kts",
    path: "settings.gradle.kts",
    language: "kotlin",
    content: `pluginManagement {
    repositories {
        google()
        mavenCentral()
        gradlePluginPortal()
    }
}
dependencyResolutionManagement {
    repositoriesMode.set(RepositoriesMode.FAIL_ON_PROJECT_REPOS)
    repositories {
        google()
        mavenCentral()
    }
}

rootProject.name = "InvestmentTracker"
include(":app")`
  },
  {
    name: "libs.versions.toml",
    path: "gradle/libs.versions.toml",
    language: "toml",
    content: `[versions]
agp = "8.2.2"
kotlin = "1.9.24"
coreKtx = "1.12.0"
junit = "4.13.2"
junitVersion = "1.1.5"
espressoCore = "3.5.1"
lifecycleRuntimeKtx = "2.7.0"
activityCompose = "1.8.2"
composeBom = "2023.08.00"

[libraries]
androidx-core-ktx = { group = "androidx.core", name = "core-ktx", value = "coreKtx" }
androidx-lifecycle-runtime-ktx = { group = "androidx.lifecycle", name = "lifecycle-runtime-ktx", value = "lifecycleRuntimeKtx" }
androidx-activity-compose = { group = "androidx.activity", name = "activity-compose", value = "activityCompose" }
androidx-compose-bom = { group = "androidx.compose", name = "compose-bom", value = "composeBom" }
androidx-compose-ui = { group = "androidx.compose.ui", name = "ui" }
androidx-compose-graphics = { group = "androidx.compose.ui", name = "ui-graphics" }
androidx-compose-ui-tooling = { group = "androidx.compose.ui", name = "ui-tooling" }
androidx-compose-ui-tooling-preview = { group = "androidx.compose.ui", name = "ui-tooling-preview" }
androidx-compose-ui-test-manifest = { group = "androidx.compose.ui", name = "ui-test-manifest" }
androidx-compose-ui-test-junit4 = { group = "androidx.compose.ui", name = "ui-test-junit4" }
androidx-compose-material3 = { group = "androidx.compose.material3", name = "material3" }
androidx-compose-material-icons-extended = { group = "androidx.compose.material", name = "material-icons-extended" }
junit = { group = "junit", name = "junit", value = "junit" }
androidx-junit = { group = "androidx.test.ext", name = "junit", value = "junitVersion" }
androidx-espresso-core = { group = "androidx.test.espresso", name = "espresso-core", value = "espressoCore" }
androidx-lifecycle-viewmodel-compose = { group = "androidx.lifecycle", name = "lifecycle-viewmodel-compose", value = "2.7.0" }

[plugins]
android-application = { id = "com.android.application", version.ref = "agp" }
kotlin-android = { id = "org.jetbrains.kotlin.android", version.ref = "kotlin" }`
  },
  {
    name: "gradle.properties",
    path: "gradle.properties",
    language: "properties",
    content: `org.gradle.jvmargs=-Xmx2048m -Dfile.encoding=UTF-8
android.useAndroidX=true
android.enableJetifier=true
kotlin.code.style=official`
  },
  {
    name: "Theme.kt",
    path: "app/src/main/java/by/invest/tracker/ui/theme/Theme.kt",
    language: "kotlin",
    content: `package by.invest.tracker.ui.theme

import android.annotation.SuppressLint
import android.app.Activity
import android.os.Build
import androidx.compose.foundation.isSystemInDarkTheme
import androidx.compose.material3.MaterialTheme
import androidx.compose.material3.darkColorScheme
import androidx.compose.material3.dynamicDarkColorScheme
import androidx.compose.material3.dynamicLightColorScheme
import androidx.compose.material3.lightColorScheme
import androidx.compose.runtime.Composable
import androidx.compose.runtime.SideEffect
import androidx.compose.ui.graphics.Color
import androidx.compose.ui.graphics.toArgb
import androidx.compose.ui.platform.LocalContext
import androidx.compose.ui.platform.LocalView
import androidx.core.view.WindowCompat

private val DarkColorScheme = darkColorScheme(
    primary = SlatePrimary,
    onPrimary = Color(0xFF0F111A),
    secondary = SlateSecondary,
    onSecondary = Color(0xFFFFFFFF),
    tertiary = SlateTertiary,
    background = SlateDarkBg,
    onBackground = Color(0xFFE2E8F0),
    surface = SlateCardBg,
    onSurface = Color(0xFFF1F5F9),
    primaryContainer = Color(0xFF064E3B),
    onPrimaryContainer = Color(0xFFD1FAE5),
    secondaryContainer = Color(0xFF1E293B),
    onSecondaryContainer = Color(0xFFE2E8F0)
)

private val LightColorScheme = lightColorScheme(
    primary = SlateLightPrimary,
    onPrimary = Color(0xFFFFFFFF),
    secondary = SlateLightSecondary,
    onSecondary = Color(0xFFFFFFFF),
    tertiary = SlateLightTertiary,
    background = SlateLightBg,
    onBackground = Color(0xFF0F172A),
    surface = SlateLightCard,
    onSurface = Color(0xFF1E293B),
    primaryContainer = Color(0xFFD1FAE5),
    onPrimaryContainer = Color(0xFF065F46),
    secondaryContainer = Color(0xFFF1F5F9),
    onSecondaryContainer = Color(0xFF1E293B)
)

@SuppressLint("NewApi")
@Composable
fun InvestmentTrackerTheme(
    darkTheme: Boolean = isSystemInDarkTheme(),
    dynamicColor: Boolean = false, // Отключаем Динамический цвет (Monet) по умолчанию, чтобы использовать наш премиальный стиль
    content: @Composable () -> Unit
) {
    val colorScheme = when {
        dynamicColor && Build.VERSION.SDK_INT >= 31 -> {
            val context = LocalContext.current
            if (darkTheme) dynamicDarkColorScheme(context) else dynamicLightColorScheme(context)
        }
        darkTheme -> DarkColorScheme
        else -> LightColorScheme
    }
    val view = LocalView.current
    if (!view.isInEditMode) {
        SideEffect {
            val window = (view.context as Activity).window
            window.statusBarColor = colorScheme.background.toArgb()
            WindowCompat.getInsetsController(window, view).isAppearanceLightStatusBars = !darkTheme
        }
    }

    MaterialTheme(
        colorScheme = colorScheme,
        typography = AppTypography,
        content = content
    )
}`
  },
  {
    name: "Color.kt",
    path: "app/src/main/java/by/invest/tracker/ui/theme/Color.kt",
    language: "kotlin",
    content: `package by.invest.tracker.ui.theme

import androidx.compose.ui.graphics.Color

// Премиальные цвета Slate Dark (Изумруд и Тёмный сланец)
val SlateDarkBg = Color(0xFF0F111A)
val SlateCardBg = Color(0xFF1C1F30)
val SlatePrimary = Color(0xFF10B981) // Изумрудно-зеленый (основные кнопки и индикаторы)
val SlateSecondary = Color(0xFF3B82F6) // Синий акцент
val SlateTertiary = Color(0xFFF59E0B) // Янтарный

// Премиальные цвета Slate Light
val SlateLightBg = Color(0xFFF8FAFC)
val SlateLightCard = Color(0xFFFFFFFF)
val SlateLightPrimary = Color(0xFF059669)
val SlateLightSecondary = Color(0xFF2563EB)
val SlateLightTertiary = Color(0xFFD97706)`
  },
  {
    name: "Type.kt",
    path: "app/src/main/java/by/invest/tracker/ui/theme/Type.kt",
    language: "kotlin",
    content: `package by.invest.tracker.ui.theme

import androidx.compose.material3.Typography
import androidx.compose.ui.text.TextStyle
import androidx.compose.ui.text.font.FontFamily
import androidx.compose.ui.text.font.FontWeight
import androidx.compose.ui.unit.sp

val AppTypography = Typography(
    bodyLarge = TextStyle(
        fontFamily = FontFamily.Default,
        fontWeight = FontWeight.Normal,
        fontSize = 16.sp,
        lineHeight = 24.sp,
        letterSpacing = 0.5.sp
    )
)`
  }
];
