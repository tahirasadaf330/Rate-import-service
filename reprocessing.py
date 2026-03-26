#!/usr/bin/env python3
"""
Reprocessing Module for Rate Import Service

This module handles reprocessing of failed email directories by:
1. Finding directories marked for reprocessing (failed status only)
2. Cleaning up metadata flags and files 
3. Resetting processing status to start fresh
4. Preserving jera_fetched flag if it was successful
"""

import json
import shutil
from pathlib import Path
from typing import Dict, List, Optional, Any
from datetime import datetime

from database import (
    get_failed_directories_for_reprocessing,
    get_processing_status,
    mark_processing_stage,
    update_reprocessing_enabled
)
from database_flag import reset_processing_flags


class ReprocessingManager:
    """Manages the reprocessing of failed email directories."""
    
    def __init__(self, attachments_root: str = "attachments"):
        self.attachments_root = Path(attachments_root)
    
    def get_directories_for_reprocessing(self) -> List[Dict[str, Any]]:
        """
        Get all directories that are eligible for reprocessing.
        This includes failed rows and manually re-flagged stale rows.
        """
        return get_failed_directories_for_reprocessing()
    
    def reset_directory_for_reprocessing(self, directory_name: str) -> bool:
        """
        Reset a directory for reprocessing by:
        1. Cleaning metadata flags (preserving jera_fetched if true)
        2. Removing cleaned files
        3. Resetting processing status flags
        4. Disabling reprocessing flag after reset
        
        Args:
            directory_name: Name of the directory to reset
            
        Returns:
            True if successful, False otherwise
        """
        try:
            directory_path = self.attachments_root / directory_name
            if not directory_path.exists():
                print(f"Directory not found: {directory_path}")
                return False
            
            # Step 1: Clean metadata
            if not self._clean_metadata_flags(directory_path):
                return False
            
            # Step 2: Remove cleaned files
            self._remove_cleaned_files(directory_path)
            
            # Step 3: Reset processing status flags
            if not reset_processing_flags(directory_name, str(self.attachments_root)):
                return False
            
            # Step 4: Disable reprocessing flag (job done)
            update_reprocessing_enabled(
                directory_name=directory_name,
                is_reprocessing_enabled=False
            )
            
            print(f"✓ Successfully reset directory for reprocessing: {directory_name}")
            return True
            
        except Exception as e:
            print(f"✗ Error resetting directory {directory_name}: {e}")
            return False
    
    def _clean_metadata_flags(self, directory_path: Path) -> bool:
        """
        Clean processing flags from metadata.json while preserving jera_fetched if true.
        
        Removes these flags:
        - date_verification_* flags  
        - jerasoft_preprocessed, need_human_eval_jerasoft
        - need_human_eval_pre, preprocessed_results, preprocess_errors
        - final_ok, attachment_stats, comparision_result
        - rate_upload_id, results_pushed
        
        Preserves:
        - jera_fetched (only if it was true)
        - All email metadata (subject, sender, etc.)
        """
        metadata_path = directory_path / "metadata.json"
        if not metadata_path.exists():
            print(f"No metadata.json found in {directory_path}")
            return False
        
        try:
            # Read current metadata
            with metadata_path.open("r", encoding="utf-8") as f:
                metadata = json.load(f)
            
            # Store the original jera_fetched value
            preserve_jera_fetched = metadata.get("jera_fetched", False)
            
            # List of keys to remove (processing flags)
            flags_to_remove = [
                # Date verification flags
                "date_verification_ingestion", 
                "date_verification_ingestion_status",
                "date_format_identified",
                
                # Jerasoft processing flags (except jera_fetched if true)
                "best_table_name",
                "human_eval_details_jerasoft",
                "need_human_eval_jerasoft", 
                "jerasoft_preprocessed",
                
                # Preprocessing flags
                "human_eval_details_pre",
                "need_human_eval_pre",
                "preprocessed_results",
                "preprocess_errors",
                
                # Final processing flags
                "final_ok",
                "attachment_stats", 
                "comparision_result",
                "rate_upload_id",
                "results_pushed",
               # Clean up old reprocessing flags
                "reprocessing_reset_at_utc"
            ]
            
            # Remove the flags
            for flag in flags_to_remove:
                metadata.pop(flag, None)
            
            # Only remove jera_fetched if it was false/null, preserve if true
            if not preserve_jera_fetched:
                metadata.pop("jera_fetched", None)
            else:
                print(f"  Preserving jera_fetched=true for {directory_path.name}")
            
            # Update processed timestamp (reuse existing field instead of creating new one)
            metadata["processed_at_utc"] = datetime.utcnow().isoformat() + "Z"
            
            # Write back cleaned metadata
            with metadata_path.open("w", encoding="utf-8") as f:
                json.dump(metadata, f, indent=2, ensure_ascii=False)
            
            print(f"  Cleaned metadata flags for {directory_path.name}")
            return True
            
        except Exception as e:
            print(f"  Error cleaning metadata for {directory_path}: {e}")
            return False
    
    def _remove_cleaned_files(self, directory_path: Path):
        """
        Remove cleaned files and comparison results to force fresh processing.
        
        Removes files matching these patterns:
        - *_cleaned.xlsx
        - *_comparison*.xlsx  
        - *_comparision*.xlsx (typo variant)
        - *_result*.xlsx
        """
        try:
            patterns = ["*_cleaned.xlsx", "*_comparison*.xlsx", "*_comparision*.xlsx", "*_result*.xlsx"]
            removed_files = []
            
            for pattern in patterns:
                for file_path in directory_path.glob(pattern):
                    if file_path.is_file():
                        file_path.unlink()
                        removed_files.append(file_path.name)
            
            if removed_files:
                print(f"  Removed cleaned files: {', '.join(removed_files)}")
            else:
                print(f"  No cleaned files found to remove in {directory_path.name}")
                
        except Exception as e:
            print(f"  Warning: Error removing cleaned files from {directory_path}: {e}")
    

    def reprocess_all_enabled_directories(self) -> Dict[str, bool]:
        """
        Find all directories marked for reprocessing and reset them.
        
        Returns:
            Dictionary mapping directory_name -> success_status
        """
        directories = self.get_directories_for_reprocessing()
        results = {}
        
        print(f"Found {len(directories)} failed directories to check for reprocessing...")
        
        for dir_info in directories:
            directory_name = dir_info["directory_name"]
            is_enabled = dir_info["is_reprocessing_enabled"]
            
            if is_enabled:
                print(f"\nReprocessing enabled directory: {directory_name}")
                success = self.reset_directory_for_reprocessing(directory_name)
                results[directory_name] = success
            else:
                print(f"Skipping {directory_name} (reprocessing not enabled)")
        
        return results


def main():
    """CLI interface for reprocessing management."""
    import sys
    
    manager = ReprocessingManager()
    
    if len(sys.argv) < 2:
        print("Usage:")
        print("  python reprocessing.py list              # List failed directories")
        print("  python reprocessing.py reset <dir_name>  # Reset specific directory")  
        print("  python reprocessing.py process_all       # Process all enabled directories")
        return
    
    command = sys.argv[1].lower()
    
    if command == "list":
        directories = manager.get_directories_for_reprocessing()
        print(f"\nFailed directories ({len(directories)} total):")
        print("-" * 80)
        for dir_info in directories:
            status = "✓ Enabled" if dir_info["is_reprocessing_enabled"] else "✗ Disabled"
            print(f"{dir_info['directory_name']} - {status}")
            print(f"  Subject: {dir_info['email_subject']}")
            print(f"  Sender: {dir_info['sender_email']}")
            print(f"  Failed at: {dir_info['updated_at']}")
            print()
    
    elif command == "reset" and len(sys.argv) > 2:
        directory_name = sys.argv[2]
        success = manager.reset_directory_for_reprocessing(directory_name)
        if success:
            print(f"✓ Directory {directory_name} has been reset for reprocessing")
        else:
            print(f"✗ Failed to reset directory {directory_name}")
    
    elif command == "process_all":
        results = manager.reprocess_all_enabled_directories()
        print(f"\nReprocessing Summary:")
        print("-" * 40)
        for directory_name, success in results.items():
            status = "✓ Success" if success else "✗ Failed"
            print(f"{directory_name}: {status}")
    
    else:
        print("Unknown command. Use 'list', 'reset <dir_name>', or 'process_all'")


if __name__ == "__main__":
    main()
