// Global variables
let stockData = [];
let categories = [];

// Initialize the application
document.addEventListener('DOMContentLoaded', function() {
    initializeApp();
});

async function initializeApp() {
    await loadCategories();
    await loadStockData();
    setupEventListeners();
}

function setupEventListeners() {
    // Add form submission
    document.getElementById('addForm').addEventListener('submit', handleAddSubmit);
    
    // Category management
    document.getElementById('addCategoryBtn').addEventListener('click', toggleNewCategory);
    document.getElementById('category').addEventListener('change', handleCategoryChange);
    document.getElementById('addProductBtn').addEventListener('click', toggleNewProduct);
    
    // Refresh buttons
    document.getElementById('refreshEdit').addEventListener('click', loadStockData);
    document.getElementById('refreshConfirm').addEventListener('click', loadStockData);
    
    // Update button in modal
    document.getElementById('updateBtn').addEventListener('click', handleUpdate);
}

async function loadCategories() {
    try {
        categories = await google.script.run.withSuccessHandler((result) => {
            const categorySelect = document.getElementById('category');
            categorySelect.innerHTML = '<option value="">Select Category</option>';
            
            result.forEach(category => {
                const option = document.createElement('option');
                option.value = category;
                option.textContent = category;
                categorySelect.appendChild(option);
            });
            
            // Load next category number for new category input
            google.script.run.withSuccessHandler((prefix) => {
                document.getElementById('newCategoryPrefix').textContent = prefix;
            }).getNextCategoryNumber();
            
            return result;
        }).getCategories();
    } catch (error) {
        showMessage('addMessage', 'Error loading categories: ' + error, 'danger');
    }
}

async function loadStockData() {
    try {
        stockData = await google.script.run.withSuccessHandler((data) => {
            populateEditTable(data);
            populateConfirmTable(data);
            return data;
        }).getSheetData();
    } catch (error) {
        showMessage('editMessage', 'Error loading data: ' + error, 'danger');
    }
}

function populateEditTable(data) {
    const tbody = document.getElementById('editTableBody');
    tbody.innerHTML = '';
    
    data.forEach((row, index) => {
        const tr = document.createElement('tr');
        tr.innerHTML = `
            <td>${row.MainOffice}</td>
            <td>${row.SubOffice}</td>
            <td>${formatDate(row.Date)}</td>
            <td>${row.Category}</td>
            <td>${row.Product}</td>
            <td>
                <span class="badge ${row.Confirm === 'Yes' ? 'bg-success' : 'bg-warning'}">
                    ${row.Confirm}
                </span>
            </td>
            <td class="table-actions">
                <button class="btn btn-sm btn-outline-primary edit-btn" data-index="${index}">
                    <i class="bi bi-pencil"></i> Edit
                </button>
            </td>
        `;
        tbody.appendChild(tr);
    });
    
    // Add event listeners to edit buttons
    document.querySelectorAll('.edit-btn').forEach(btn => {
        btn.addEventListener('click', handleEdit);
    });
}

function populateConfirmTable(data) {
    const tbody = document.getElementById('confirmTableBody');
    tbody.innerHTML = '';
    
    // Sort data: "No" first, then "Yes"
    const sortedData = [...data].sort((a, b) => {
        if (a.Confirm === 'No' && b.Confirm === 'Yes') return -1;
        if (a.Confirm === 'Yes' && b.Confirm === 'No') return 1;
        return 0;
    });
    
    sortedData.forEach((row, index) => {
        const originalIndex = data.findIndex(item => 
            item.MainOffice === row.MainOffice &&
            item.SubOffice === row.SubOffice &&
            item.Category === row.Category &&
            item.Product === row.Product
        );
        
        const tr = document.createElement('tr');
        tr.innerHTML = `
            <td>${row.MainOffice}</td>
            <td>${row.SubOffice}</td>
            <td>${formatDate(row.Date)}</td>
            <td>${row.Category}</td>
            <td>${row.Product}</td>
            <td>
                <span class="badge ${row.Confirm === 'Yes' ? 'bg-success' : 'bg-warning'}">
                    ${row.Confirm}
                </span>
            </td>
            <td class="table-actions">
                ${row.Confirm === 'No' ? 
                    `<button class="btn btn-sm btn-success confirm-btn" data-index="${originalIndex}">
                        <i class="bi bi-check"></i> Confirm
                    </button>` :
                    `<button class="btn btn-sm btn-warning reverse-btn" data-index="${originalIndex}">
                        <i class="bi bi-arrow-counterclockwise"></i> Reverse
                    </button>`
                }
            </td>
        `;
        tbody.appendChild(tr);
    });
    
    // Add event listeners to action buttons
    document.querySelectorAll('.confirm-btn').forEach(btn => {
        btn.addEventListener('click', handleConfirm);
    });
    
    document.querySelectorAll('.reverse-btn').forEach(btn => {
        btn.addEventListener('click', handleReverse);
    });
}

function toggleNewCategory() {
    const newCategorySection = document.getElementById('newCategorySection');
    const categorySelect = document.getElementById('category');
    const addProductBtn = document.getElementById('addProductBtn');
    
    if (newCategorySection.style.display === 'none' || !newCategorySection.style.display) {
        newCategorySection.style.display = 'block';
        categorySelect.disabled = true;
        addProductBtn.disabled = false;
        
        // Clear and focus new category input
        document.getElementById('newCategory').value = '';
        document.getElementById('newCategory').focus();
    } else {
        newCategorySection.style.display = 'none';
        categorySelect.disabled = false;
        addProductBtn.disabled = true;
        categorySelect.focus();
    }
}

function toggleNewProduct() {
    const newProductSection = document.getElementById('newProductSection');
    const productInput = document.getElementById('product');
    const category = getSelectedCategory();
    
    if (!category) {
        showMessage('addMessage', 'Please select or add a category first', 'warning');
        return;
    }
    
    if (newProductSection.style.display === 'none' || !newProductSection.style.display) {
        newProductSection.style.display = 'block';
        productInput.disabled = true;
        
        // Get next product number for the selected category
        google.script.run.withSuccessHandler((prefix) => {
            document.getElementById('newProductPrefix').textContent = prefix;
        }).getNextProductNumber(category);
        
        // Clear and focus new product input
        document.getElementById('newProduct').value = '';
        document.getElementById('newProduct').focus();
    } else {
        newProductSection.style.display = 'none';
        productInput.disabled = false;
        productInput.focus();
    }
}

function handleCategoryChange() {
    const category = document.getElementById('category').value;
    const addProductBtn = document.getElementById('addProductBtn');
    
    if (category) {
        addProductBtn.disabled = false;
        
        // Update product prefix based on selected category
        google.script.run.withSuccessHandler((prefix) => {
            document.getElementById('productPrefix').textContent = prefix;
        }).getNextProductNumber(category);
    } else {
        addProductBtn.disabled = true;
    }
}

function getSelectedCategory() {
    const newCategory = document.getElementById('newCategory').value.trim();
    const selectedCategory = document.getElementById('category').value;
    
    if (document.getElementById('newCategorySection').style.display === 'block' && newCategory) {
        const prefix = document.getElementById('newCategoryPrefix').textContent;
        return prefix + newCategory;
    }
    
    return selectedCategory;
}

function getSelectedProduct() {
    const newProduct = document.getElementById('newProduct').value.trim();
    const product = document.getElementById('product').value.trim();
    
    if (document.getElementById('newProductSection').style.display === 'block' && newProduct) {
        const prefix = document.getElementById('newProductPrefix').textContent;
        return prefix + newProduct;
    }
    
    const categoryPrefix = document.getElementById('productPrefix').textContent;
    return categoryPrefix + product;
}

async function handleAddSubmit(e) {
    e.preventDefault();
    
    const mainOffice = document.getElementById('mainOffice').value;
    const subOffice = document.getElementById('subOffice').value;
    const date = document.getElementById('date').value;
    const category = getSelectedCategory();
    const product = getSelectedProduct();
    
    // Validation
    if (!mainOffice || !subOffice || !date || !category || !product) {
        showMessage('addMessage', 'Please fill all required fields', 'warning');
        return;
    }
    
    // Date validation
    if (!isValidDate(date)) {
        showMessage('addMessage', 'Please enter a valid date in DD/MM/YYYY format', 'warning');
        return;
    }
    
    try {
        const result = await google.script.run.withSuccessHandler((response) => {
            if (response.success) {
                showMessage('addMessage', response.message, 'success');
                document.getElementById('addForm').reset();
                resetCategoryAndProduct();
                loadStockData();
                loadCategories();
            } else {
                showMessage('addMessage', response.message, 'danger');
            }
        }).withFailureHandler((error) => {
            showMessage('addMessage', 'Error: ' + error, 'danger');
        }).addStockItem(mainOffice, subOffice, date, category, product);
    } catch (error) {
        showMessage('addMessage', 'Error: ' + error, 'danger');
    }
}

function handleEdit(e) {
    const index = e.target.closest('.edit-btn').dataset.index;
    const row = stockData[index];
    
    // Populate modal fields
    document.getElementById('editRowIndex').value = index;
    document.getElementById('editMainOffice').value = row.MainOffice;
    document.getElementById('editSubOffice').value = row.SubOffice;
    document.getElementById('editDate').value = formatDate(row.Date);
    document.getElementById('editCategory').value = row.Category;
    document.getElementById('editProduct').value = row.Product;
    document.getElementById('editConfirm').value = row.Confirm;
    
    // Clear previous messages
    document.getElementById('editMessage').innerHTML = '';
    
    // Show modal
    const modal = new bootstrap.Modal(document.getElementById('editModal'));
    modal.show();
}

async function handleUpdate() {
    const index = document.getElementById('editRowIndex').value;
    const mainOffice = document.getElementById('editMainOffice').value;
    const subOffice = document.getElementById('editSubOffice').value;
    const date = document.getElementById('editDate').value;
    const category = document.getElementById('editCategory').value;
    const product = document.getElementById('editProduct').value;
    const confirm = document.getElementById('editConfirm').value;
    
    // Validation
    if (!mainOffice || !subOffice || !date || !category || !product) {
        showMessage('editMessage', 'Please fill all required fields', 'warning');
        return;
    }
    
    if (!isValidDate(date)) {
        showMessage('editMessage', 'Please enter a valid date in DD/MM/YYYY format', 'warning');
        return;
    }
    
    try {
        const result = await google.script.run.withSuccessHandler((response) => {
            if (response.success) {
                showMessage('editMessage', response.message, 'success');
                loadStockData();
                
                // Close modal after successful update
                setTimeout(() => {
                    bootstrap.Modal.getInstance(document.getElementById('editModal')).hide();
                }, 1500);
            } else {
                showMessage('editMessage', response.message, 'danger');
            }
        }).updateStockItem(index, mainOffice, subOffice, date, category, product, confirm);
    } catch (error) {
        showMessage('editMessage', 'Error: ' + error, 'danger');
    }
}

async function handleConfirm(e) {
    const index = e.target.closest('.confirm-btn').dataset.index;
    
    try {
        const result = await google.script.run.withSuccessHandler((response) => {
            if (response.success) {
                loadStockData();
            } else {
                alert('Error: ' + response.message);
            }
        }).confirmItem(index, 'Yes');
    } catch (error) {
        alert('Error: ' + error);
    }
}

async function handleReverse(e) {
    const index = e.target.closest('.reverse-btn').dataset.index;
    
    try {
        const result = await google.script.run.withSuccessHandler((response) => {
            if (response.success) {
                loadStockData();
            } else {
                alert('Error: ' + response.message);
            }
        }).confirmItem(index, 'No');
    } catch (error) {
        alert('Error: ' + error);
    }
}

function resetCategoryAndProduct() {
    // Reset category section
    document.getElementById('newCategorySection').style.display = 'none';
    document.getElementById('category').disabled = false;
    document.getElementById('category').value = '';
    document.getElementById('newCategory').value = '';
    
    // Reset product section
    document.getElementById('newProductSection').style.display = 'none';
    document.getElementById('product').disabled = false;
    document.getElementById('product').value = '';
    document.getElementById('newProduct').value = '';
    document.getElementById('productPrefix').textContent = '1-';
    document.getElementById('addProductBtn').disabled = true;
}

function isValidDate(dateString) {
    const regex = /^\d{2}\/\d{2}\/\d{4}$/;
    if (!regex.test(dateString)) return false;
    
    const parts = dateString.split('/');
    const day = parseInt(parts[0], 10);
    const month = parseInt(parts[1], 10);
    const year = parseInt(parts[2], 10);
    
    if (year < 1000 || year > 3000 || month === 0 || month > 12) return false;
    
    const monthLength = [31, 28, 31, 30, 31, 30, 31, 31, 30, 31, 30, 31];
    
    // Adjust for leap years
    if (year % 400 === 0 || (year % 100 !== 0 && year % 4 === 0)) {
        monthLength[1] = 29;
    }
    
    return day > 0 && day <= monthLength[month - 1];
}

function formatDate(dateString) {
    if (!dateString) return '';
    
    const date = new Date(dateString);
    if (isNaN(date.getTime())) return dateString;
    
    const day = String(date.getDate()).padStart(2, '0');
    const month = String(date.getMonth() + 1).padStart(2, '0');
    const year = date.getFullYear();
    
    return `${day}/${month}/${year}`;
}

function showMessage(elementId, message, type) {
    const element = document.getElementById(elementId);
    element.innerHTML = `
        <div class="alert alert-${type} alert-dismissible fade show">
            ${message}
            <button type="button" class="btn-close" data-bs-dismiss="alert"></button>
        </div>
    `;
}
