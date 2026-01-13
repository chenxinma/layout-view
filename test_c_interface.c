#include <stdio.h>
#include <stdlib.h>
#include <dlfcn.h>
#include <string.h>

int main() {
    // Load the dynamic library
    void* handle = dlopen("./target/release/liblayout_view.so", RTLD_LAZY);
    if (!handle) {
        fprintf(stderr, "Error: %s\n", dlerror());
        return 1;
    }

    // Define function types
    typedef char* (*classify_func_t)(const char*);
    typedef void (*free_func_t)(char*);

    // Get the function from the library
    classify_func_t classify_excel_sheets_c = (classify_func_t) dlsym(handle, "classify_excel_sheets_c");
    free_func_t free_c_string = (free_func_t) dlsym(handle, "free_c_string");

    const char* error;
    if ((error = dlerror()) != NULL) {
        fprintf(stderr, "Error: %s\n", error);
        return 1;
    }

    // Call the function
    const char* file_path = "./files/test_data.xlsx";
    printf("Classifying file: %s\n", file_path);
    
    char* result = classify_excel_sheets_c(file_path);
    if (result != NULL) {
        printf("Result: %s\n", result);
        free_c_string(result);
    } else {
        printf("Function returned NULL\n");
    }

    // Close the library
    dlclose(handle);
    return 0;
}