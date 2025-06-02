import pandas as pd
import re
from datetime import datetime

class WeiboDataProcessor:
    """微博数据处理器"""
    
    def __init__(self):
        pass
    
    def get_basic_stats(self, df):
        """获取基础统计信息"""
        if df.empty:
            return {}
        
        stats = {
            '总微博数': len(df),
            '总转发数': df['转发数'].sum(),
            '总评论数': df['评论数'].sum(), 
            '总点赞数': df['点赞数'].sum(),
            '平均转发数': round(df['转发数'].mean(), 2),
            '平均评论数': round(df['评论数'].mean(), 2),
            '平均点赞数': round(df['点赞数'].mean(), 2),
            '最热微博转发数': df['转发数'].max(),
            '最热微博评论数': df['评论数'].max(),
            '最热微博点赞数': df['点赞数'].max()
        }
        return stats
    
    def get_time_distribution(self, df):
        """获取时间分布数据"""
        if df.empty:
            return None
        
        try:
            # 过滤有效时间数据
            valid_time_df = df[df['发布时间'] != 'N/A'].copy()
            if valid_time_df.empty:
                return None
            
            # 转换时间格式
            valid_time_df['发布时间'] = pd.to_datetime(valid_time_df['发布时间'], errors='coerce')
            valid_time_df = valid_time_df.dropna(subset=['发布时间'])
            
            if valid_time_df.empty:
                return None
            
            # 按小时统计
            valid_time_df['小时'] = valid_time_df['发布时间'].dt.hour
            hourly_counts = valid_time_df['小时'].value_counts().sort_index()
            
            return hourly_counts
        except Exception as e:
            print(f"处理时间分布数据时出错: {str(e)}")
            return None
    
    def get_author_stats(self, df):
        """获取作者统计数据"""
        if df.empty:
            return None
        
        try:
            author_stats = df.groupby('微博作者').agg({
                '微博id': 'count',
                '转发数': 'sum',
                '评论数': 'sum',
                '点赞数': 'sum'
            }).rename(columns={'微博id': '微博数量'})
            
            # 按微博数量排序，取前10
            top_authors = author_stats.sort_values('微博数量', ascending=False).head(10)
            return top_authors
        except Exception as e:
            print(f"处理作者统计数据时出错: {str(e)}")
            return None
    
    def get_content_length_stats(self, df):
        """获取内容长度统计"""
        if df.empty:
            return None
        
        try:
            df['内容长度'] = df['微博内容'].astype(str).apply(len)
            length_stats = {
                '平均长度': round(df['内容长度'].mean(), 2),
                '最长内容': df['内容长度'].max(),
                '最短内容': df['内容长度'].min(),
                '中位数长度': df['内容长度'].median()
            }
            return length_stats
        except Exception as e:
            print(f"处理内容长度统计时出错: {str(e)}")
            return None
    
    def clean_content(self, content):
        """清理微博内容"""
        if pd.isna(content) or content == 'N/A':
            return ''
        
        # 移除HTML标签
        content = re.sub(r'<[^>]+>', '', str(content))
        # 移除多余空格
        content = re.sub(r'\s+', ' ', content).strip()
        return content
    
    def export_to_excel(self, df, filename):
        """导出数据到Excel"""
        try:
            with pd.ExcelWriter(filename, engine='xlsxwriter') as writer:
                # 写入主要数据
                df.to_excel(writer, sheet_name='微博数据', index=False)
                
                # 写入统计信息
                stats = self.get_basic_stats(df)
                stats_df = pd.DataFrame(list(stats.items()), columns=['指标', '数值'])
                stats_df.to_excel(writer, sheet_name='统计信息', index=False)
                
                # 写入作者统计
                author_stats = self.get_author_stats(df)
                if author_stats is not None:
                    author_stats.to_excel(writer, sheet_name='作者统计')
                
            return True
        except Exception as e:
            print(f"导出Excel时出错: {str(e)}")
            return False 