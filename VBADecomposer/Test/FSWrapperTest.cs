/*
 * Unit-test for utility File System Wrapper class
 * User: nergal-perm
 * Date: 31.07.2015
 * Time: 11:42
 * 
 */
using System;
using NUnit.Framework;
using VBADecomposer.Main;

namespace VBADecomposer.Test {
	[TestFixture]
	public class FSWrapperTest {
		private FSWrapperFake _target;
		
		[SetUp]
		public void setUp() {
			_target = new FSWrapperFake();
		}
		
		[Test]
		public void ShouldCheckForWorkbookExistence() {
			Assert.IsFalse(_target.WorkbookExists());
			
			_target.WorkbookPath = @"C:\existing_workbook.xls";
			Assert.IsTrue(_target.WorkbookExists());
			
			_target.WorkbookPath = @"C:\non_existing_workbook.xls";
			Assert.IsFalse(_target.WorkbookExists());
			
			_target.WorkbookPath = @"C:\any_workbook.xls";
			Assert.IsFalse(_target.WorkbookExists());
		}
	}
}
