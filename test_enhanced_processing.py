"""
Test script for enhanced document processing features
"""

import os
import sys
from enhanced_document_processor import EnhancedDocumentProcessor

def test_enhanced_processor():
    """Test the enhanced document processor with sample data"""
    print("🧪 Testing Enhanced Document Processor...")
    
    try:
        # Initialize processor
        processor = EnhancedDocumentProcessor(fuzzy_threshold=85)
        print("✅ Enhanced processor initialized successfully")
        
        # Test fuzzy matching
        test_text = """
        The CHD performance must be 14 days or more for the adjustment to apply.
        If CHD exceeds threshold, 2.5 bps per day will be deducted from gross incentive.
        CV Tier 1 volume calculations should include Lo-ROC adjustments.
        NACV includes Bus Unadjusted minus AR 180 Days and Writeoff reserves.
        """
        
        print("\n📝 Testing Fuzzy Matching...")
        
        # Test CHD keyword matching
        chd_matches = processor.fuzzy_match_keywords(test_text, processor.chd_keywords)
        print(f"CHD matches found: {len(chd_matches)}")
        for match in chd_matches[:3]:  # Show first 3
            print(f"  - '{match.match}' (score: {match.score:.1f}) in context: '{match.context[:50]}...'")
        
        # Test Excel pattern matching  
        excel_matches = processor.fuzzy_match_keywords(test_text, processor.excel_patterns)
        print(f"\nExcel pattern matches found: {len(excel_matches)}")
        for match in excel_matches[:3]:  # Show first 3
            print(f"  - '{match.match}' (score: {match.score:.1f}) in context: '{match.context[:50]}...'")
        
        print("\n📊 Testing Table Structure Recognition...")
        
        # Test with existing files if available
        sample_files = [
            'input file 2.xlsx',
            'input file 3.xlsx', 
            'Sample 2 contract.docx',
            'Sample 3.docx'
        ]
        
        for filename in sample_files:
            filepath = os.path.join(os.path.dirname(__file__), filename)
            if os.path.exists(filepath):
                print(f"\nTesting with: {filename}")
                
                if filename.endswith('.xlsx'):
                    # Test enhanced Excel detection
                    result = processor.enhanced_excel_detection(filepath)
                    print(f"  Detection type: {result['detected_type']}")
                    print(f"  Confidence: {result['confidence']:.2f}")
                    print(f"  Reasoning: {result['reasoning']}")
                    
                elif filename.endswith('.docx'):
                    # Test enhanced CHD extraction
                    chd_result = processor.extract_context_aware_chd_rules(filepath)
                    print(f"  CHD confidence: {chd_result.get('confidence', 0):.2f}")
                    print(f"  Matches found: {chd_result.get('matches_found', 0)}")
                    
                    # Test enhanced BIR extraction
                    bir_result = processor.extract_enhanced_bir_table(filepath)
                    print(f"  BIR entries found: {len(bir_result) if bir_result else 0}")
            else:
                print(f"  ⚠️  File not found: {filename}")
        
        print("\n✅ Enhanced processing tests completed!")
        
    except ImportError as e:
        print(f"❌ Import error: {e}")
        print("Install required packages: pip install -r requirements.txt")
        return False
    except Exception as e:
        print(f"❌ Test failed: {e}")
        return False
        
    return True

def test_integration():
    """Test integration with main app"""
    print("\n🔗 Testing Integration with Main App...")
    
    try:
        from app import ENHANCED_PROCESSING, enhanced_processor
        
        if ENHANCED_PROCESSING and enhanced_processor:
            print("✅ Enhanced processing integrated successfully")
            print(f"   Fuzzy threshold: {enhanced_processor.fuzzy_threshold}")
            print(f"   Available methods: {len([m for m in dir(enhanced_processor) if not m.startswith('_')])}")
        else:
            print("⚠️  Enhanced processing not available in main app")
            
    except Exception as e:
        print(f"❌ Integration test failed: {e}")
        return False
        
    return True

if __name__ == "__main__":
    print("🚀 Enhanced Document Processing Test Suite")
    print("=" * 50)
    
    success = True
    success &= test_enhanced_processor()
    success &= test_integration()
    
    print("\n" + "=" * 50)
    if success:
        print("🎉 All tests passed!")
    else:
        print("❌ Some tests failed. Check the output above.")
        sys.exit(1)