"""
Multiprocessing Search Test
"""

import os
import sys
from multiprocessing import cpu_count, freeze_support

def run_test():
    print("=" * 60)
    print("Multiprocessing Search Test")
    print("=" * 60)

    # CPU cores
    print(f"\nCPU cores: {cpu_count()}")

    # Process count to use
    num_processes = min(cpu_count(), 4)
    print(f"Will use: {num_processes} processes")

    # Test search module
    try:
        from search_print import search_order_in_pdf, search_order_in_folder
        print("\n[OK] Search module loaded")
    except Exception as e:
        print(f"\n[FAIL] Search module load failed: {e}")
        return False

    # Test multiprocessing
    try:
        from multiprocessing import Pool
        
        def test_function(x):
            return x * x
        
        with Pool(processes=2) as pool:
            results = pool.map(test_function, [1, 2, 3, 4])
        
        print(f"[OK] Multiprocessing test passed: {results}")
    except Exception as e:
        print(f"[FAIL] Multiprocessing test failed: {e}")
        return False

    print("\n" + "=" * 60)
    print("All tests passed!")
    print("=" * 60)
    
    return True


if __name__ == "__main__":
    freeze_support()
    
    success = run_test()
    
    if success:
        print("\nMultiprocessing is working!")
        print("\nHow to test in main program:")
        print("1. Run: python main.py")
        print("2. Go to 'Search/Print' tab")
        print("3. Select folder with 10+ PDFs")
        print("4. Enter order number")
        print("5. Click 'Search'")
        print("\nLook for log message:")
        print("  'N processes parallel search started...'")
